using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SQLite;

namespace Safety_System
{
    public static class DataManager
    {
        private const string ConfigFile = "sys_config.txt";
        private const string DbFileName = "SafetyData.sqlite";
        public static string BasePath { get; private set; } = AppDomain.CurrentDomain.BaseDirectory;

        private static string GetConnString()
        {
            return string.Format("Data Source={0};Version=3;", Path.Combine(BasePath, DbFileName));
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

        // 檢查資料庫檔案是否存在
        public static bool IsDbFileExists()
        {
            return File.Exists(Path.Combine(BasePath, DbFileName));
        }

        // 建立並初始化水處理資料表
        public static void CreateWaterTable()
        {
            using (var conn = new SQLiteConnection(GetConnString()))
            {
                conn.Open();
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
            }
        }

        // 動態新增欄位
        public static void AddColumnToWaterTable(string columnName)
        {
            using (var conn = new SQLiteConnection(GetConnString()))
            {
                conn.Open();
                // 注意：SQLite 新增欄位需使用 ALTER TABLE
                string sql = string.Format("ALTER TABLE WaterMeterReadings ADD COLUMN [{0}] TEXT", columnName);
                using (var cmd = new SQLiteCommand(sql, conn)) { cmd.ExecuteNonQuery(); }
            }
        }

        // 取得資料 (區間查詢或最近30筆)
        public static DataTable GetWaterData(string start, string end)
        {
            DataTable dt = new DataTable();
            using (var conn = new SQLiteConnection(GetConnString()))
            {
                conn.Open();
                string sql = "SELECT * FROM WaterMeterReadings WHERE LogDate BETWEEN @s AND @e ORDER BY LogDate DESC";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@s", start);
                    cmd.Parameters.AddWithValue("@e", end);
                    using (var adapter = new SQLiteDataAdapter(cmd)) { adapter.Fill(dt); }
                }

                // 如果沒資料或日期錯誤，改抓最近30筆
                if (dt.Rows.Count == 0)
                {
                    string sqlTop30 = "SELECT * FROM WaterMeterReadings ORDER BY LogDate DESC LIMIT 30";
                    using (var cmd30 = new SQLiteCommand(sqlTop30, conn))
                    using (var adapter30 = new SQLiteDataAdapter(cmd30)) { dt.Clear(); adapter30.Fill(dt); }
                }
            }
            return dt;
        }

        // 即時更新單一欄位數據
        public static void UpdateWaterCell(int id, string columnName, object value)
        {
            using (var conn = new SQLiteConnection(GetConnString()))
            {
                conn.Open();
                string sql = string.Format("UPDATE WaterMeterReadings SET [{0}] = @val WHERE Id = @id", columnName);
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@val", value.ToString());
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
