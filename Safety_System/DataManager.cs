using System;
using System.IO;
using System.Text;
using System.Data.SQLite; // 必須引用此命名空間

namespace Safety_System
{
    public static class DataManager
    {
        private const string ConfigFile = "sys_config.txt";
        private const string DbFileName = "SafetyData.sqlite";
        
        public static string BasePath { get; private set; } = AppDomain.CurrentDomain.BaseDirectory;

        private static string GetConnString()
        {
            string dbPath = Path.Combine(BasePath, DbFileName);
            return string.Format("Data Source={0};Version=3;", dbPath);
        }

        public static void LoadConfig()
        {
            if (File.Exists(ConfigFile))
            {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (Directory.Exists(savedPath)) BasePath = savedPath;
            }
            InitializeDatabase();
        }

        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            File.WriteAllText(ConfigFile, newPath, Encoding.UTF8);
            InitializeDatabase();
        }

        private static void InitializeDatabase()
        {
            using (var conn = new SQLiteConnection(GetConnString()))
            {
                conn.Open();
                string sql = @"
                    CREATE TABLE IF NOT EXISTS Inspection (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        LogDate TEXT,
                        Location TEXT,
                        Inspector TEXT,
                        Status TEXT
                    );";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            using (var conn = new SQLiteConnection(GetConnString()))
            {
                conn.Open();
                string sql = "INSERT INTO Inspection (LogDate, Location, Inspector, Status) VALUES (@date, @loc, @ins, @sta)";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@loc", location);
                    cmd.Parameters.AddWithValue("@ins", inspector);
                    cmd.Parameters.AddWithValue("@sta", status);
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
