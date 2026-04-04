using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SQLite;
using System.Threading;

namespace Safety_System
{
    public static class DataManager
    {
        private const string ConfigFile = "sys_config.txt";
        
        // 預設路徑鎖定在程式目錄下的 DB 資料夾
        public static string BasePath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");

        static DataManager() {
            // 🟢 確保程式啟動時 DB 資料夾即存在
            EnsureDirectoryExists(BasePath);
        }

        private static void EnsureDirectoryExists(string path) {
            try {
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            } catch { }
        }

        private static string GetConnString(string dbName) {
            EnsureDirectoryExists(BasePath);
            string fullPath = Path.Combine(BasePath, dbName + ".sqlite");
            return string.Format("Data Source={0};Version=3;Default Timeout=15;Pooling=True;Max Pool Size=100;", fullPath);
        }

        public static void LoadConfig() {
            if (File.Exists(ConfigFile)) {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (Directory.Exists(savedPath)) BasePath = savedPath;
            }
            EnsureDirectoryExists(BasePath);
        }

        public static void SetBasePath(string newPath) {
            BasePath = newPath;
            File.WriteAllText(ConfigFile, newPath, Encoding.UTF8);
            EnsureDirectoryExists(BasePath);
        }

        private static void ExecuteWithRetry(string dbName, Action<SQLiteConnection> dbAction) {
            int maxRetries = 5;
            for (int i = 1; i <= maxRetries; i++) {
                try {
                    using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                        conn.Open();
                        using (var pragma = new SQLiteCommand("PRAGMA journal_mode=WAL;", conn)) { pragma.ExecuteNonQuery(); }
                        dbAction(conn); return;
                    }
                } catch (SQLiteException) { if (i == maxRetries) throw; Thread.Sleep(500); }
            }
        }

        public static void InitTable(string dbName, string tableName, string schemaSql) {
            ExecuteWithRetry(dbName, conn => { using (var cmd = new SQLiteCommand(schemaSql, conn)) { cmd.ExecuteNonQuery(); } });
        }

        public static DataTable GetTableData(string dbName, string tableName, string dateCol, string start, string end) {
            DataTable dt = new DataTable();
            ExecuteWithRetry(dbName, conn => {
                string sql = string.Format("SELECT * FROM [{0}] WHERE [{1}] BETWEEN @s AND @e ORDER BY [{1}] DESC", tableName, dateCol);
                using (var cmd = new SQLiteCommand(sql, conn)) {
                    cmd.Parameters.AddWithValue("@s", start); cmd.Parameters.AddWithValue("@e", end);
                    using (var adp = new SQLiteDataAdapter(cmd)) { adp.Fill(dt); }
                }
            });
            return dt;
        }

        public static void UpsertRecord(string dbName, string tableName, DataRow row) {
            ExecuteWithRetry(dbName, conn => {
                if (row.Table.Columns.Contains("日期") && row["日期"] != DBNull.Value) {
                    if (DateTime.TryParse(row["日期"].ToString(), out DateTime dt)) row["日期"] = dt.ToString("yyyy-MM-dd");
                }
                bool isUpdate = row["Id"] != DBNull.Value;
                StringBuilder sb = new StringBuilder();
                if (isUpdate) {
                    sb.Append(string.Format("UPDATE [{0}] SET ", tableName));
                    foreach (DataColumn col in row.Table.Columns) if (col.ColumnName != "Id") sb.Append(string.Format("[{0}]=@{0}, ", col.ColumnName));
                    sb.Remove(sb.Length - 2, 2).Append(" WHERE Id=@Id");
                } else {
                    sb.Append(string.Format("INSERT INTO [{0}] (", tableName));
                    foreach (DataColumn col in row.Table.Columns) if (col.ColumnName != "Id") sb.Append(string.Format("[{0}], ", col.ColumnName));
                    sb.Remove(sb.Length - 2, 2).Append(") VALUES (");
                    foreach (DataColumn col in row.Table.Columns) if (col.ColumnName != "Id") sb.Append(string.Format("@{0}, ", col.ColumnName));
                    sb.Remove(sb.Length - 2, 2).Append(")");
                }
                using (var cmd = new SQLiteCommand(sb.ToString(), conn)) {
                    foreach (DataColumn col in row.Table.Columns) cmd.Parameters.AddWithValue("@" + col.ColumnName, row[col.ColumnName] ?? "");
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public static void DeleteRecord(string dbName, string tableName, int id) {
            ExecuteWithRetry(dbName, conn => {
                using (var cmd = new SQLiteCommand(string.Format("DELETE FROM [{0}] WHERE Id=@Id", tableName), conn)) {
                    cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery();
                }
            });
        }

        public static void AddColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand(string.Format("ALTER TABLE [{0}] ADD COLUMN [{1}] TEXT", tableName, colName), conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void RenameColumn(string dbName, string tableName, string oldN, string newN) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand(string.Format("ALTER TABLE [{0}] RENAME COLUMN [{1}] TO [{2}]", tableName, oldN, newN), conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand(string.Format("ALTER TABLE [{0}] DROP COLUMN [{1}]", tableName, colName), conn)) { cmd.ExecuteNonQuery(); }
        });
    }
}
