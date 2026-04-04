using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Threading;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string ConfigFile = Path.Combine(AppDir, "sys_config.txt");
        private static readonly string KeyConfigFile = Path.Combine(AppDir, "table_keys.txt"); // 存放防重寫設定

        public static string BasePath { get; private set; }

        public static void LoadConfig()
        {
            // 🟢 需求 3.2：預設路徑為程式目錄下的 DB 子資料夾
            BasePath = Path.Combine(AppDir, "DB");

            if (File.Exists(ConfigFile))
            {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (!string.IsNullOrEmpty(savedPath)) BasePath = savedPath;
            }

            // 確保資料夾存在以防異常
            if (!Directory.Exists(BasePath)) Directory.CreateDirectory(BasePath);
        }

        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            File.WriteAllText(ConfigFile, newPath, Encoding.UTF8);
            if (!Directory.Exists(BasePath)) Directory.CreateDirectory(BasePath);
        }

        private static string GetConnString(string dbName)
        {
            if (!Directory.Exists(BasePath)) Directory.CreateDirectory(BasePath);
            string fullPath = Path.Combine(BasePath, dbName + ".sqlite");
            return string.Format("Data Source={0};Version=3;Default Timeout=15;Pooling=True;Max Pool Size=100;", fullPath);
        }

        // 🟢 [讀取防重寫設定] (純文字 TXT，| 分隔)
        public static (string col1, string col2) GetTableKeys(string dbName, string tableName)
        {
            if (!File.Exists(KeyConfigFile)) return ("", "");
            foreach (var line in File.ReadAllLines(KeyConfigFile, Encoding.UTF8))
            {
                var p = line.Split('|');
                if (p.Length >= 4 && p[0] == dbName && p[1] == tableName) return (p[2], p[3]);
            }
            return ("", "");
        }

        // 🟢 [儲存防重寫設定]
        public static void SaveTableKeys(string dbName, string tableName, string col1, string col2)
        {
            List<string> lines = new List<string>();
            if (File.Exists(KeyConfigFile)) lines = new List<string>(File.ReadAllLines(KeyConfigFile, Encoding.UTF8));
            
            lines.RemoveAll(x => x.StartsWith($"{dbName}|{tableName}|")); // 移除舊設定
            lines.Add($"{dbName}|{tableName}|{col1}|{col2}"); // 寫入新設定
            File.WriteAllLines(KeyConfigFile, lines, Encoding.UTF8);
        }

        // 取得資料表的所有欄位 (供設定介面使用)
        public static List<string> GetColumnNames(string dbName, string tableName)
        {
            var cols = new List<string>();
            try {
                ExecuteWithRetry(dbName, conn => {
                    using (var cmd = new SQLiteCommand($"PRAGMA table_info([{tableName}])", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) cols.Add(r["name"].ToString()); }
                });
            } catch { /* 忽略尚未建立的表 */ }
            return cols;
        }

        // === 原有 DML/DDL 簡化保留 ===
        public static void InitTable(string dbName, string tableName, string createSql) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand(createSql, conn)) cmd.ExecuteNonQuery();
        });

        public static DataTable GetTableData(string dbName, string tableName, string dateCol, string start, string end)
        {
            DataTable dt = new DataTable();
            ExecuteWithRetry(dbName, conn => {
                string q = string.IsNullOrEmpty(start) ? $"SELECT * FROM [{tableName}]" : $"SELECT * FROM [{tableName}] WHERE [{dateCol}] BETWEEN @s AND @e";
                using (var cmd = new SQLiteCommand(q, conn)) {
                    if (!string.IsNullOrEmpty(start)) { cmd.Parameters.AddWithValue("@s", start); cmd.Parameters.AddWithValue("@e", end); }
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            });
            return dt;
        }

        // 🟢 寫入與更新核心邏輯
        public static void UpsertRecord(string dbName, string tableName, DataRow row)
        {
            // 🟢 需求 2：日期強制統一格式處理
            foreach (DataColumn col in row.Table.Columns)
            {
                if (col.ColumnName.Contains("日期") && row[col] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col].ToString()))
                {
                    if (DateTime.TryParse(row[col].ToString(), out DateTime dt))
                        row[col] = dt.ToString("yyyy-MM-dd"); // 強制轉為 yyyy-MM-dd
                }
            }

            // 🟢 需求 4：讀取設定，判斷是否觸發防重寫 (複寫邏輯)
            var keys = GetTableKeys(dbName, tableName);
            if (!string.IsNullOrEmpty(keys.col1) && row.Table.Columns.Contains(keys.col1))
            {
                string v1 = row[keys.col1]?.ToString();
                string v2 = (!string.IsNullOrEmpty(keys.col2) && row.Table.Columns.Contains(keys.col2)) ? row[keys.col2]?.ToString() : null;

                string query = $"SELECT Id FROM [{tableName}] WHERE [{keys.col1}] = @v1";
                if (!string.IsNullOrEmpty(keys.col2)) query += $" AND [{keys.col2}] = @v2";

                int existingId = -1;
                ExecuteWithRetry(dbName, conn => {
                    using (var cmd = new SQLiteCommand(query, conn)) {
                        cmd.Parameters.AddWithValue("@v1", v1 ?? "");
                        if (!string.IsNullOrEmpty(keys.col2)) cmd.Parameters.AddWithValue("@v2", v2 ?? "");
                        var res = cmd.ExecuteScalar();
                        if (res != null && res != DBNull.Value) existingId = Convert.ToInt32(res);
                    }
                });

                // 如果找到重複的鍵值，強制將當前 Row 的 Id 覆寫，讓後續邏輯視為 UPDATE
                if (existingId != -1)
                {
                    bool isReadOnly = row.Table.Columns["Id"].ReadOnly;
                    row.Table.Columns["Id"].ReadOnly = false;
                    row["Id"] = existingId;
                    row.Table.Columns["Id"].ReadOnly = isReadOnly;
                }
            }

            // 執行實際的 INSERT 或 UPDATE
            ExecuteWithRetry(dbName, conn => {
                bool isUpdate = row.Table.Columns.Contains("Id") && row["Id"] != DBNull.Value && !string.IsNullOrEmpty(row["Id"].ToString()) && Convert.ToInt32(row["Id"]) > 0;
                var cmd = new SQLiteCommand(conn);
                string sql = "";

                if (isUpdate) {
                    sql = $"UPDATE [{tableName}] SET ";
                    List<string> sets = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        sets.Add($"[{col.ColumnName}]=@{col.ColumnName}");
                        cmd.Parameters.AddWithValue("@" + col.ColumnName, row[col] ?? "");
                    }
                    sql += string.Join(", ", sets) + " WHERE Id=@Id";
                    cmd.Parameters.AddWithValue("@Id", row["Id"]);
                } else {
                    List<string> cols = new List<string>(), vals = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        cols.Add($"[{col.ColumnName}]");
                        vals.Add($"@{col.ColumnName}");
                        cmd.Parameters.AddWithValue("@" + col.ColumnName, row[col] ?? "");
                    }
                    sql = $"INSERT INTO [{tableName}] ({string.Join(", ", cols)}) VALUES ({string.Join(", ", vals)})";
                }
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
            });
        }

        private static void ExecuteWithRetry(string dbName, Action<SQLiteConnection> dbAction)
        {
            int maxRetries = 5;
            for (int i = 1; i <= maxRetries; i++) {
                try {
                    using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                        conn.Open(); dbAction(conn); return;
                    }
                } catch (SQLiteException ex) when (ex.ResultCode == SQLiteErrorCode.Locked || ex.ResultCode == SQLiteErrorCode.Busy) {
                    if (i == maxRetries) throw;
                    Thread.Sleep(100 * i);
                }
            }
        }
    }
}
