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

        // 智慧重試封裝
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
                        using (var pragmaCmd = new SQLiteCommand("PRAGMA journal_mode=WAL;", conn))
                        {
                            pragmaCmd.ExecuteNonQuery();
                        }
                        dbAction(conn);
                        return;
                    }
                }
                catch (SQLiteException ex)
                {
                    if (ex.ResultCode == SQLiteErrorCode.Busy || ex.ResultCode == SQLiteErrorCode.Locked)
                    {
                        if (attempt == maxRetries)
                            throw new Exception("多人同時存取導致資料庫忙碌。\n系統已自動重試 5 次依然失敗，請稍後再試！");
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
                        LogDate TEXT,
                        WasteWaterDischarge TEXT,
                        WasteWaterInflow TEXT,
                        Recycle6Inch TEXT,
                        DualMediaA TEXT,
                        DualMediaB TEXT,
                        StorageTank TEXT,
                        SoftWaterA TEXT,
                        SoftWaterB TEXT,
                        SoftWaterC TEXT
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

        // 單格即時更新
        public static void UpdateWaterCell(int id, string columnName, object value)
        {
            try
            {
                ExecuteWithRetry(conn => {
                    string sql = string.Format("UPDATE WaterMeterReadings SET [{0}] = @val WHERE Id = @id", columnName);
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@val", value?.ToString() ?? "");
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "資料儲存異常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 智慧儲存 (複寫或新增)
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

        // 🟢 新增：刪除指定資料列
        public static void DeleteWaterRecord(int id)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show("刪除資料失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static DataTable GetWaterData(string start, string end)
        {
            DataTable dt = new DataTable();
            try
            {
                using (var conn = new SQLiteConnection(GetConnString()))
                {
                    conn.Open();
                    using (var pragmaCmd = new SQLiteCommand("PRAGMA journal_mode=WAL;", conn)) { pragmaCmd.ExecuteNonQuery(); }

                    string sql = "SELECT * FROM WaterMeterReadings WHERE LogDate BETWEEN @s AND @e ORDER BY LogDate DESC";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@s", start);
                        cmd.Parameters.AddWithValue("@e", end);
                        using (var adapter = new SQLiteDataAdapter(cmd)) { adapter.Fill(dt); }
                    }

                    if (dt.Rows.Count == 0)
                    {
                        string sqlTop30 = "SELECT * FROM WaterMeterReadings ORDER BY LogDate DESC LIMIT 30";
                        using (var cmd30 = new SQLiteCommand(sqlTop30, conn))
                        using (var adapter30 = new SQLiteDataAdapter(cmd30)) { dt.Clear(); adapter30.Fill(dt); }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取資料庫失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return dt;
        }

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            try
            {
                ExecuteWithRetry(conn => {
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "巡檢紀錄儲存異常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
