using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public static class BackupManager
    {
        public static string BackupPath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB_Backups");
        public static int KeepCount { get; private set; } = 4; 
        public static DateTime LastBackupDate { get; private set; } = DateTime.MinValue;

        public static void LoadConfig()
        {
            string dbPath = DataManager.GetSysSetting("BackupPath", "");
            string dbKeep = DataManager.GetSysSetting("BackupKeepCount", "");
            string dbDate = DataManager.GetSysSetting("LastBackupDate", "");

            if (!string.IsNullOrEmpty(dbPath)) BackupPath = dbPath;
            if (int.TryParse(dbKeep, out int count)) KeepCount = count;
            if (DateTime.TryParse(dbDate, out DateTime date)) LastBackupDate = date;

            if (!Directory.Exists(BackupPath)) Directory.CreateDirectory(BackupPath);
        }

        public static void SaveConfig(string path, int count)
        {
            BackupPath = path;
            KeepCount = count;
            DataManager.SetSysSetting("BackupPath", path);
            DataManager.SetSysSetting("BackupKeepCount", count.ToString());
            UpdateConfigDate(LastBackupDate);
            if (!Directory.Exists(BackupPath)) Directory.CreateDirectory(BackupPath);
        }

        private static void UpdateConfigDate(DateTime date)
        {
            LastBackupDate = date;
            DataManager.SetSysSetting("LastBackupDate", date.ToString("yyyy-MM-dd"));
        }

        public static void RunAutoBackup()
        {
            LoadConfig();
            if ((DateTime.Today - LastBackupDate).TotalDays >= 7)
            {
                ExecuteBackup();
            }
        }

        public static void ExecuteBackup()
        {
            try
            {
                if (!Directory.Exists(DataManager.BasePath)) return;

                string folderName = DateTime.Now.ToString("yyyyMMdd_HHmm");
                string targetDir = Path.Combine(BackupPath, folderName);
                Directory.CreateDirectory(targetDir);

                // 🟢 修復：改用 SQLite 原生備份 API (Hot Backup)，徹底解決 File.Copy 造成的檔案鎖死與衝突問題
                string[] files = Directory.GetFiles(DataManager.BasePath, "*.sqlite");
                foreach (var file in files)
                {
                    string destFile = Path.Combine(targetDir, Path.GetFileName(file));
                    SafeBackupSQLite(file, destFile);
                }

                string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                if (File.Exists(sysDbPath))
                {
                    SafeBackupSQLite(sysDbPath, Path.Combine(targetDir, "SystemConfig.sqlite"));
                }

                UpdateConfigDate(DateTime.Today);
                CleanupOldBackups();
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料庫備份失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 🟢 原生安全備份方法
        private static void SafeBackupSQLite(string sourcePath, string destPath)
        {
            try
            {
                using (var source = new SQLiteConnection($"Data Source={sourcePath};Version=3;"))
                using (var destination = new SQLiteConnection($"Data Source={destPath};Version=3;"))
                {
                    source.Open();
                    destination.Open();
                    source.BackupDatabase(destination, "main", "main", -1, null, 0);
                }
            }
            catch (Exception ex)
            {
                // 若原生備份失敗 (極少數情況)，降級嘗試 File.Copy
                try { File.Copy(sourcePath, destPath, true); } catch { throw ex; }
            }
        }

        private static void CleanupOldBackups()
        {
            try
            {
                var directories = new DirectoryInfo(BackupPath).GetDirectories()
                                    .OrderByDescending(d => d.CreationTime)
                                    .ToList();

                if (directories.Count > KeepCount)
                {
                    for (int i = KeepCount; i < directories.Count; i++)
                    {
                        directories[i].Delete(true);
                    }
                }
            }
            catch { }
        }
    }
}
