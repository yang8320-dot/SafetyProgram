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

        // 🟢 防呆機制 1: 連線字串加入 Default Timeout=15 與 Pooling，讓 SQLite 原生具備排隊等待能力
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

        // ==========================================
        // 🟢 防呆機制 2: 建立全域的「智慧重試」封裝方法
        // 所有寫入動作都透過此方法，遇到鎖定時會自動重新嘗試
        // ==========================================
        private static void ExecuteWithRetry(Action<SQLiteConnection> dbAction)
        {
            int maxRetries = 5;      // 最多重試 5 次
            int delayMs = 500;       // 每次失敗等待 0.5 秒

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var conn = new SQLiteConnection(GetConnString()))
                    {
                        conn.Open();
                        
                        // 🟢 防呆機制 3: 每次連線都確保開啟 WAL 模式，讓讀寫不互卡
                        using (var pragmaCmd = new SQLiteCommand("PRAGMA journal_mode=WAL;", conn))
                        {
                            pragmaCmd.ExecuteNonQuery();
                        }

                        dbAction(conn); // 執行真正的資料庫操作
                        return;         // 成功則跳出迴圈
                    }
                }
                catch (SQLiteException ex)
                {
                    // 偵測到錯誤碼 5 (Busy) 或 6 (Locked) 代表別人正在寫入
                    if (ex.ResultCode == SQLiteErrorCode.Busy || ex.ResultCode == SQLiteErrorCode.Locked)
                    {
                        if (attempt == maxRetries)
                        {
                            throw new Exception("多人同時存取導致資料庫忙碌。\n系統已自動重試 5 次依然失敗，請稍後再試！");
                        }
                        Thread.Sleep(delayMs); // 暫停一下再試
                    }
                    else
                    {
                        throw; // 若是其他錯誤 (如 SQL 語法錯) 則直接拋出
                    }
                }
            }
        }

        // ==========================================
        // 實作各項資料庫操作 (全面改用 ExecuteWithRetry)
        // ==========================================

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

        // 🟢 即時存檔的寫入點，加入了最外層的防呆保護，避免應用程式崩潰
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
                // 如果真的重試 5 次都失敗，或是網路斷線，只會彈出警告，不會讓程式閃退
                MessageBox.Show(ex.Message, "資料儲存異常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 讀取因為不會鎖死資料庫，所以保持原本的寫法即可，但同樣加入 WAL 確保讀取順暢
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

        // 巡檢紀錄也順便套用新的防呆寫入法
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
