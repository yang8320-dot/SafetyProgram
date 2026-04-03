using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SQLite;
using System.Threading;
using System.Windows.Forms;

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

        public static void LoadConfig()
        {
            if (File.Exists(ConfigFile))
            {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (Directory.Exists(savedPath)) BasePath = savedPath;
            }
        }

        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            File.WriteAllText(ConfigFile, newPath, Encoding.UTF8);
        }

        public static bool IsDbFileExists()
        {
            return File.Exists(Path.Combine(BasePath, DbFileName));
        }

        private static void ExecuteWithRetry(Action<SQLiteConnection> dbAction)
        {
            int maxRetries = 5;
            int delayMs = 500;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var conn = new SQLiteConnection(GetConnString()))
                    {
                        conn.Open();
                        using (var pragmaCmd = new SQLiteCommand("PRAGMA journal_mode=WAL;", conn)) { pragmaCmd.ExecuteNonQuery(); }
                        dbAction(conn);
                        return;
                    }
                }
                catch (SQLiteException ex)
                {
                    if (ex.ResultCode == SQLiteErrorCode.Busy || ex.ResultCode == SQLiteErrorCode.Locked)
                    {
                        if (attempt == maxRetries) throw new Exception("多人同時存取導致資料庫忙碌。\n系統已自動重試 5 次依然失敗，請稍後再試！");
                        Thread.Sleep(delayMs);
                    }
                    else throw;
                }
            }
        }

        public static void CreateWaterTable()
        {
            ExecuteWithRetry(conn => {
                string sql = @"
                    CREATE TABLE IF NOT EXISTS WaterMeterReadings (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        [日期] TEXT,
                        [廢水處理] TEXT,
                        [廢水進流] TEXT,
                        [納管排放] TEXT,
                        [回收6吋] TEXT,
                        [雙介質A] TEXT,
                        [雙介質B] TEXT,
                        [貯存池] TEXT,
                        [軟水A] TEXT,
                        [軟水B] TEXT,
                        [軟水C] TEXT
                    );";
                using (var cmd = new SQLiteCommand(sql, conn)) { cmd.ExecuteNonQuery(); }
            });
        }

        public static void AddColumnToWaterTable(string columnName)
        {
            ExecuteWithRetry(conn => {
                string sql = string.Format("ALTER TABLE WaterMeterReadings ADD COLUMN [{0}] TEXT", columnName);
                using (var cmd = new SQLiteCommand(sql, conn)) { cmd.ExecuteNonQuery(); }
            });
        }

        public static void UpsertWaterRecord(DataRow row)
        {
            ExecuteWithRetry(conn => {
                bool isUpdate = row["Id"] != DBNull.Value;
                string sql;

                if (isUpdate)
                {
                    StringBuilder sb = new StringBuilder("UPDATE WaterMeterReadings SET ");
                    foreach (DataColumn col in row.Table.Columns)
                    {
                        if (col.ColumnName == "Id") continue;
                        sb.Append(string.Format("[{0}] = @{0}, ", col.ColumnName));
                    }
                    sb.Remove(sb.Length - 2, 2);
                    sb.Append(" WHERE Id = @Id");
                    sql = sb.ToString();
                }
                else
                {
                    StringBuilder sbCols = new StringBuilder("INSERT INTO WaterMeterReadings (");
                    StringBuilder sbVals = new StringBuilder("VALUES (");
                    foreach (DataColumn col in row.Table.Columns)
                    {
                        if (col.ColumnName == "Id") continue;
                        sbCols.Append(string.Format("[{0}], ", col.ColumnName));
                        sbVals.Append(string.Format("@{0}, ", col.ColumnName));
                    }
                    sbCols.Remove(sbCols.Length - 2, 2).Append(") ");
                    sbVals.Remove(sbVals.Length - 2, 2).Append(") ");
                    sql = sbCols.ToString() + sbVals.ToString();
                }

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    foreach (DataColumn col in row.Table.Columns)
                    {
                        cmd.Parameters.AddWithValue("@" + col.ColumnName, row[col.ColumnName] ?? "");
                    }
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public static void DeleteWaterRecord(int id)
        {
            ExecuteWithRetry(conn => {
                string sql = "DELETE FROM WaterMeterReadings WHERE Id = @Id";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.ExecuteNonQuery();
                }
            });
        }

        // 🟢 強制校正資料庫中的錯亂日期格式
        private static void AutoFixDateFormats()
        {
            ExecuteWithRetry(conn => {
                DataTable dt = new DataTable();
                // 找出格式不對的日期 (有斜線或長度不是10碼)
                using (var cmd = new SQLiteCommand("SELECT Id, [日期] FROM WaterMeterReadings WHERE [日期] LIKE '%/%' OR length([日期]) <> 10", conn))
                {
                    using (var adapter = new SQLiteDataAdapter(cmd)) { adapter.Fill(dt); }
                }

                foreach (DataRow row in dt.Rows)
                {
                    if (row["日期"] != DBNull.Value && DateTime.TryParse(row["日期"].ToString(), out DateTime d))
                    {
                        using (var updateCmd = new SQLiteCommand("UPDATE WaterMeterReadings SET [日期] = @d WHERE Id = @id", conn))
                        {
                            updateCmd.Parameters.AddWithValue("@d", d.ToString("yyyy-MM-dd"));
                            updateCmd.Parameters.AddWithValue("@id", row["Id"]);
                            updateCmd.ExecuteNonQuery();
                        }
                    }
                }
            });
        }

        public static DataTable GetWaterData(string start, string end)
        {
            // 在讀取之前先背景洗一次資料，確保比對準確無誤
            AutoFixDateFormats();

            DataTable dt = new DataTable();
            try
            {
                using (var conn = new SQLiteConnection(GetConnString()))
                {
                    conn.Open();
                    string sql = "SELECT * FROM WaterMeterReadings WHERE [日期] BETWEEN @s AND @e ORDER BY [日期] DESC";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@s", start);
                        cmd.Parameters.AddWithValue("@e", end);
                        using (var adapter = new SQLiteDataAdapter(cmd)) { adapter.Fill(dt); }
                    }

                    // 如果真的區間內沒資料，才會抓最近的30筆
                    if (dt.Rows.Count == 0)
                    {
                        string sqlTop30 = "SELECT * FROM WaterMeterReadings ORDER BY [日期] DESC LIMIT 30";
                        using (var cmd30 = new SQLiteCommand(sqlTop30, conn))
                        using (var adapter30 = new SQLiteDataAdapter(cmd30)) { dt.Clear(); adapter30.Fill(dt); }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("讀取失敗：" + ex.Message); }
            return dt;
        }

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            try
            {
                ExecuteWithRetry(conn => {
                    string createSql = @"
                        CREATE TABLE IF NOT EXISTS Inspection (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            LogDate TEXT,
                            Location TEXT,
                            Inspector TEXT,
                            Status TEXT
                        );";
                    using (var createCmd = new SQLiteCommand(createSql, conn)) { createCmd.ExecuteNonQuery(); }

                    string sql = "INSERT INTO Inspection (LogDate, Location, Inspector, Status) VALUES (@date, @loc, @ins, @sta)";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@date", date);
                        cmd.Parameters.AddWithValue("@loc", location);
                        cmd.Parameters.AddWithValue("@ins", inspector);
                        cmd.Parameters.AddWithValue("@sta", status);
                        cmd.ExecuteNonQuery();
                    }
                });
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "儲存異常"); }
        }
    }
}
