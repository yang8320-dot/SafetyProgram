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
        private const string DbFileName = "SafetyData.sqlite";
        public static string BasePath { get; private set; } = AppDomain.CurrentDomain.BaseDirectory;

        private static string GetConnString()
        {
            return string.Format("Data Source={0};Version=3;Default Timeout=15;Pooling=True;Max Pool Size=100;", Path.Combine(BasePath, DbFileName));
        }

        // 🟢 補回：讀取設定檔路徑
        public static void LoadConfig()
        {
            if (File.Exists(ConfigFile))
            {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (Directory.Exists(savedPath)) BasePath = savedPath;
            }
        }

        // 🟢 補回：設定資料庫路徑
        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            File.WriteAllText(ConfigFile, newPath, Encoding.UTF8);
        }

        // 🟢 補回：判斷資料庫檔案是否存在
        public static bool IsDbFileExists()
        {
            return File.Exists(Path.Combine(BasePath, DbFileName));
        }

        private static void ExecuteWithRetry(Action<SQLiteConnection> dbAction)
        {
            int maxRetries = 5;
            for (int i = 1; i <= maxRetries; i++)
            {
                try
                {
                    using (var conn = new SQLiteConnection(GetConnString()))
                    {
                        conn.Open();
                        using (var pragma = new SQLiteCommand("PRAGMA journal_mode=WAL;", conn)) { pragma.ExecuteNonQuery(); }
                        dbAction(conn); return;
                    }
                }
                catch (SQLiteException) { if (i == maxRetries) throw; Thread.Sleep(500); }
            }
        }

        public static void InitTable(string tableName, string schemaSql)
        {
            ExecuteWithRetry(conn => {
                using (var cmd = new SQLiteCommand(schemaSql, conn)) { cmd.ExecuteNonQuery(); }
            });
        }

        public static DataTable GetTableData(string tableName, string dateCol, string start, string end)
        {
            DataTable dt = new DataTable();
            ExecuteWithRetry(conn => {
                string sql = string.Format("SELECT * FROM [{0}] WHERE [{1}] BETWEEN @s AND @e ORDER BY [{1}] DESC", tableName, dateCol);
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@s", start);
                    cmd.Parameters.AddWithValue("@e", end);
                    using (var adp = new SQLiteDataAdapter(cmd)) { adp.Fill(dt); }
                }
            });
            return dt;
        }

        public static void UpsertRecord(string tableName, DataRow row)
        {
            ExecuteWithRetry(conn => {
                bool isUpdate = row["Id"] != DBNull.Value;
                StringBuilder sb = new StringBuilder();
                if (isUpdate) {
                    sb.Append(string.Format("UPDATE [{0}] SET ", tableName));
                    foreach (DataColumn col in row.Table.Columns) 
                        if (col.ColumnName != "Id") sb.Append(string.Format("[{0}]=@{0}, ", col.ColumnName));
                    sb.Remove(sb.Length - 2, 2).Append(" WHERE Id=@Id");
                } else {
                    sb.Append(string.Format("INSERT INTO [{0}] (", tableName));
                    foreach (DataColumn col in row.Table.Columns) 
                        if (col.ColumnName != "Id") sb.Append(string.Format("[{0}], ", col.ColumnName));
                    sb.Remove(sb.Length - 2, 2).Append(") VALUES (");
                    foreach (DataColumn col in row.Table.Columns) 
                        if (col.ColumnName != "Id") sb.Append(string.Format("@{0}, ", col.ColumnName));
                    sb.Remove(sb.Length - 2, 2).Append(")");
                }
                using (var cmd = new SQLiteCommand(sb.ToString(), conn)) {
                    foreach (DataColumn col in row.Table.Columns) cmd.Parameters.AddWithValue("@" + col.ColumnName, row[col.ColumnName] ?? "");
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public static void DeleteRecord(string tableName, int id)
        {
            ExecuteWithRetry(conn => {
                using (var cmd = new SQLiteCommand(string.Format("DELETE FROM [{0}] WHERE Id=@Id", tableName), conn)) {
                    cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery();
                }
            });
        }

        public static void AddColumn(string tableName, string colName) => ExecuteWithRetry(conn => {
            using (var cmd = new SQLiteCommand(string.Format("ALTER TABLE [{0}] ADD COLUMN [{1}] TEXT", tableName, colName), conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void RenameColumn(string tableName, string oldN, string newN) => ExecuteWithRetry(conn => {
            using (var cmd = new SQLiteCommand(string.Format("ALTER TABLE [{0}] RENAME COLUMN [{1}] TO [{2}]", tableName, oldN, newN), conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropColumn(string tableName, string colName) => ExecuteWithRetry(conn => {
            using (var cmd = new SQLiteCommand(string.Format("ALTER TABLE [{0}] DROP COLUMN [{1}]", tableName, colName), conn)) { cmd.ExecuteNonQuery(); }
        });
    }
}
