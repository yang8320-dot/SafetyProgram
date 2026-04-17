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
            string fullPath = Path.Combine(BasePath, dbName + ".sqlite");
            return string.Format("Data Source={0};Version=3;Default Timeout=30;Pooling=True;Max Pool Size=100;", fullPath);
        }

        private static void ExecuteWithRetry(string dbName, Action<SQLiteConnection> dbAction)
        {
            int maxRetries = 5;
            for (int i = 1; i <= maxRetries; i++) {
                try {
                    using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                        conn.Open(); 
                        dbAction(conn); 
                        return;
                    }
                } catch (SQLiteException ex) when (ex.ResultCode == SQLiteErrorCode.Locked || ex.ResultCode == SQLiteErrorCode.Busy) {
                    if (i == maxRetries) throw;
                    Thread.Sleep(100 * i);
                }
            }
        }

        // 🚀 極速批次儲存方法 (支援事務處理，並升級 4 欄位交叉比對與 NULL 防呆)
        public static bool BulkSaveTable(string dbName, string tableName, DataTable dt)
        {
            try {
                using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        var keys = GetTableKeys(dbName, tableName);

                        // 收集有設定的判斷欄位
                        List<string> activeKeys = new List<string>();
                        if (!string.IsNullOrEmpty(keys.col1) && dt.Columns.Contains(keys.col1)) activeKeys.Add(keys.col1);
                        if (!string.IsNullOrEmpty(keys.col2) && dt.Columns.Contains(keys.col2)) activeKeys.Add(keys.col2);
                        if (!string.IsNullOrEmpty(keys.col3) && dt.Columns.Contains(keys.col3)) activeKeys.Add(keys.col3);
                        if (!string.IsNullOrEmpty(keys.col4) && dt.Columns.Contains(keys.col4)) activeKeys.Add(keys.col4);

                        foreach (DataRow row in dt.Rows) {
                            if (row.RowState == DataRowState.Deleted) continue;

                            // 處理日期格式化
                            foreach (DataColumn col in dt.Columns) {
                                if (col.ColumnName.Contains("日期") && row[col] != DBNull.Value) {
                                    if (DateTime.TryParse(row[col].ToString(), out DateTime d))
                                        row[col] = d.ToString("yyyy-MM-dd");
                                }
                            }

                            int existingId = -1;
                            
                            // 🟢 判斷重複邏輯 (動態組合 1~4 個判斷條件)
                            if (activeKeys.Count > 0) {
                                List<string> whereClauses = new List<string>();
                                foreach (var k in activeKeys) {
                                    string safeKey = k.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                    // 使用 IFNULL 完美解決 Excel 匯入時 空字串與 NULL 不等價 造成的防呆失效問題
                                    whereClauses.Add($"IFNULL([{k}], '') = IFNULL(@{safeKey}, '')");
                                }

                                string qCheck = $"SELECT Id FROM [{tableName}] WHERE " + string.Join(" AND ", whereClauses);

                                using (var cmdCheck = new SQLiteCommand(qCheck, conn, trans)) {
                                    foreach (var k in activeKeys) {
                                        string safeKey = k.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                        // 🟢 加入 (object) 強制轉型，解決 C# 7.3 編譯錯誤
                                        object val = row[k] != DBNull.Value ? (object)row[k].ToString().Trim() : DBNull.Value;
                                        cmdCheck.Parameters.AddWithValue("@" + safeKey, val);
                                    }
                                    
                                    var res = cmdCheck.ExecuteScalar();
                                    if (res != null && res != DBNull.Value) existingId = Convert.ToInt32(res);
                                }
                            }

                            // 執行 Upsert (Insert or Update)
                            bool isUpdate = (existingId != -1) || (dt.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0);
                            using (var cmd = new SQLiteCommand(conn)) {
                                cmd.Transaction = trans;
                                List<string> sqlParts = new List<string>();
                                
                                if (isUpdate) {
                                    int targetId = existingId != -1 ? existingId : Convert.ToInt32(row["Id"]);
                                    string sql = $"UPDATE [{tableName}] SET ";
                                    foreach (DataColumn col in dt.Columns) {
                                        if (col.ColumnName == "Id") continue;
                                        
                                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                        sqlParts.Add($"[{col.ColumnName}]=@{safeParamName}");
                                        // 🟢 加入 (object) 強制轉型
                                        object val = row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value;
                                        cmd.Parameters.AddWithValue("@" + safeParamName, val);
                                    }
                                    cmd.CommandText = sql + string.Join(", ", sqlParts) + " WHERE Id=" + targetId;
                                } else {
                                    List<string> colNames = new List<string>();
                                    List<string> paramNames = new List<string>();
                                    foreach (DataColumn col in dt.Columns) {
                                        if (col.ColumnName == "Id") continue;
                                        
                                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                        colNames.Add($"[{col.ColumnName}]");
                                        paramNames.Add($"@{safeParamName}");
                                        // 🟢 加入 (object) 強制轉型
                                        object val = row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value;
                                        cmd.Parameters.AddWithValue("@" + safeParamName, val);
                                    }
                                    cmd.CommandText = $"INSERT INTO [{tableName}] ({string.Join(", ", colNames)}) VALUES ({string.Join(", ", paramNames)})";
                                }
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                return true;
            } catch (Exception ex) {
                MessageBox.Show("批次儲存失敗：" + ex.Message);
                return false;
            }
        }

        public static (string col1, string col2, string col3, string col4) GetTableKeys(string dbName, string tableName)
        {
            if (!File.Exists(KeyConfigFile)) return ("", "", "", "");
            foreach (var line in File.ReadAllLines(KeyConfigFile, Encoding.UTF8)) {
                var p = line.Split('|');
                if (p.Length >= 4 && p[0] == dbName && p[1] == tableName) {
                    string c1 = p[2];
                    string c2 = p[3];
                    string c3 = p.Length >= 5 ? p[4] : "";
                    string c4 = p.Length >= 6 ? p[5] : "";
                    return (c1, c2, c3, c4);
                }
            }
            return ("", "", "", "");
        }

        public static void SaveTableKeys(string dbName, string tableName, string col1, string col2, string col3, string col4)
        {
            List<string> lines = new List<string>();
            if (File.Exists(KeyConfigFile)) lines = new List<string>(File.ReadAllLines(KeyConfigFile, Encoding.UTF8));
            lines.RemoveAll(x => x.StartsWith($"{dbName}|{tableName}|")); 
            lines.Add($"{dbName}|{tableName}|{col1}|{col2}|{col3}|{col4}"); 
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

        public static DataTable GetLatestRecords(string dbName, string tableName, int limit = 30)
        {
            DataTable dt = new DataTable();
            ExecuteWithRetry(dbName, conn => {
                string q = $"SELECT * FROM (SELECT * FROM [{tableName}] ORDER BY Id DESC LIMIT @limit) sub ORDER BY Id ASC";
                using (var cmd = new SQLiteCommand(q, conn)) {
                    cmd.Parameters.AddWithValue("@limit", limit);
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            });
            return dt;
        }

        public static bool ValidateAndSaveTable(string dbName, string tableName, DataTable dt)
        {
            return BulkSaveTable(dbName, tableName, dt);
        }

        public static void UpsertRecord(string dbName, string tableName, DataRow row)
        {
            ExecuteWithRetry(dbName, conn => {
                bool isUpdate = row.Table.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0;
                var cmd = new SQLiteCommand(conn);
                if (isUpdate) {
                    List<string> sets = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                        sets.Add($"[{col.ColumnName}]=@{safeParamName}");
                        // 🟢 加入 (object) 強制轉型
                        cmd.Parameters.AddWithValue("@" + safeParamName, row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value);
                    }
                    cmd.CommandText = $"UPDATE [{tableName}] SET {string.Join(", ", sets)} WHERE Id=" + row["Id"];
                } else {
                    List<string> c = new List<string>();
                    List<string> v = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                        c.Add($"[{col.ColumnName}]"); 
                        v.Add($"@{safeParamName}");
                        // 🟢 加入 (object) 強制轉型
                        cmd.Parameters.AddWithValue("@" + safeParamName, row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value);
                    }
                    cmd.CommandText = $"INSERT INTO [{tableName}] ({string.Join(", ", c)}) VALUES ({string.Join(", ", v)})";
                }
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

        public static void DropTable(string dbName, string tableName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"DROP TABLE IF EXISTS [{tableName}]", conn)) { cmd.ExecuteNonQuery(); }
        });
    }
}
