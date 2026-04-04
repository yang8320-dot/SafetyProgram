using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string ConfigFile = Path.Combine(AppDir, "sys_config.txt");
        private static readonly string KeyConfigFile = Path.Combine(AppDir, "table_keys.txt"); 

        public static string BasePath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");

        public static void LoadConfig()
        {
            BasePath = Path.Combine(AppDir, "DB");
            if (File.Exists(ConfigFile))
            {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (!string.IsNullOrEmpty(savedPath)) BasePath = savedPath;
            }
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

        public static void SaveTableKeys(string dbName, string tableName, string col1, string col2)
        {
            List<string> lines = new List<string>();
            if (File.Exists(KeyConfigFile)) lines = new List<string>(File.ReadAllLines(KeyConfigFile, Encoding.UTF8));
            
            lines.RemoveAll(x => x.StartsWith($"{dbName}|{tableName}|")); 
            lines.Add($"{dbName}|{tableName}|{col1}|{col2}"); 
            File.WriteAllLines(KeyConfigFile, lines, Encoding.UTF8);
        }

        public static List<string> GetColumnNames(string dbName, string tableName)
        {
            var cols = new List<string>();
            try {
                ExecuteWithRetry(dbName, conn => {
                    using (var cmd = new SQLiteCommand($"PRAGMA table_info([{tableName}])", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) cols.Add(r["name"].ToString()); }
                });
            } catch { } 
            return cols;
        }

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

        public static bool ValidateAndSaveTable(string dbName, string tableName, DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                if (row.RowState == DataRowState.Deleted) continue;

                foreach (DataColumn col in dt.Columns)
                {
                    if (col.ColumnName.Contains("日期") && row[col] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        if (!DateTime.TryParse(row[col].ToString(), out DateTime date))
                        {
                            MessageBox.Show(
                                $"【存檔中斷】\n\n第 {i + 1} 列的【{col.ColumnName}】輸入了無效的格式：「{row[col]}」\n\n請修正為正確的日期 (例如: 2024-04-05) 後，再重新按儲存。", 
                                "日期格式錯誤", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Warning);
                            return false; 
                        }
                    }
                }
            }

            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState != DataRowState.Deleted)
                {
                    UpsertRecord(dbName, tableName, row);
                }
            }
            return true;
        }

        // 🟢 核心修改區：比對差異並彈出確認視窗
        public static void UpsertRecord(string dbName, string tableName, DataRow row)
        {
            foreach (DataColumn col in row.Table.Columns) {
                if (col.ColumnName.Contains("日期") && row[col] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col].ToString())) {
                    if (DateTime.TryParse(row[col].ToString(), out DateTime dt))
                        row[col] = dt.ToString("yyyy-MM-dd"); 
                }
            }

            var keys = GetTableKeys(dbName, tableName);
            int existingId = -1;

            if (!string.IsNullOrEmpty(keys.col1) && row.Table.Columns.Contains(keys.col1))
            {
                string v1 = row[keys.col1]?.ToString();
                string v2 = (!string.IsNullOrEmpty(keys.col2) && row.Table.Columns.Contains(keys.col2)) ? row[keys.col2]?.ToString() : null;

                string query = $"SELECT Id FROM [{tableName}] WHERE [{keys.col1}] = @v1";
                if (!string.IsNullOrEmpty(keys.col2)) query += $" AND [{keys.col2}] = @v2";

                ExecuteWithRetry(dbName, conn => {
                    using (var cmd = new SQLiteCommand(query, conn)) {
                        cmd.Parameters.AddWithValue("@v1", v1 ?? "");
                        if (!string.IsNullOrEmpty(keys.col2)) cmd.Parameters.AddWithValue("@v2", v2 ?? "");
                        var res = cmd.ExecuteScalar();
                        if (res != null && res != DBNull.Value) existingId = Convert.ToInt32(res);
                    }
                });
            }

            // 🟢 若發現重複資料，進行新舊資料比對
            if (existingId != -1) {
                DataTable oldDt = new DataTable();
                ExecuteWithRetry(dbName, conn => {
                    using (var cmd = new SQLiteCommand($"SELECT * FROM [{tableName}] WHERE Id=@Id", conn)) {
                        cmd.Parameters.AddWithValue("@Id", existingId);
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(oldDt);
                    }
                });

                if (oldDt.Rows.Count > 0) {
                    DataRow oldRow = oldDt.Rows[0];
                    StringBuilder msg = new StringBuilder();
                    msg.AppendLine($"發現已存在的重複紀錄 (依據防重寫規則)，是否確定要覆蓋此紀錄？\n");
                    
                    bool hasDifferences = false;

                    // 逐一比對欄位，只列出有被修改的內容
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        
                        string oldVal = oldRow.Table.Columns.Contains(col.ColumnName) ? oldRow[col.ColumnName]?.ToString() : "";
                        string newVal = row[col]?.ToString() ?? "";

                        if (oldVal != newVal) {
                            msg.AppendLine($"【{col.ColumnName}】");
                            msg.AppendLine($"  原本值：{(string.IsNullOrEmpty(oldVal) ? "(空)" : oldVal)}");
                            msg.AppendLine($"  取代值：{(string.IsNullOrEmpty(newVal) ? "(空)" : newVal)}\n");
                            hasDifferences = true;
                        }
                    }

                    // 如果資料完全沒有變動，直接略過，不吵使用者也不重複寫入
                    if (!hasDifferences) {
                        return; 
                    }

                    // 彈出確認視窗
                    var result = MessageBox.Show(msg.ToString(), "確認覆蓋取代", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    
                    // 如果使用者選「否」，直接 return 中斷這一筆的存檔
                    if (result == DialogResult.No) {
                        return; 
                    }
                }

                // 若選擇「是」，將舊資料的 Id 指定給新列，觸發覆蓋(UPDATE)邏輯
                bool isReadOnly = row.Table.Columns["Id"].ReadOnly;
                row.Table.Columns["Id"].ReadOnly = false;
                row["Id"] = existingId;
                row.Table.Columns["Id"].ReadOnly = isReadOnly;
            }

            // 執行寫入或更新
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

        public static void DeleteRecord(string dbName, string tableName, int id) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"DELETE FROM [{tableName}] WHERE Id=@Id", conn)) { cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery(); }
        });

        public static void AddColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] ADD COLUMN [{colName}] TEXT", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void RenameColumn(string dbName, string tableName, string oldN, string newN) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] RENAME COLUMN [{oldN}] TO [{newN}]", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] DROP COLUMN [{colName}]", conn)) { cmd.ExecuteNonQuery(); }
        });
    }
}
