/// FILE: Safety_System/BackupManager.cs ///
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
        public static int KeepCount { get; private set; } = 30; // 🟢 預設改為保留 30 份 (約一個月)
        public static int BackupIntervalDays { get; private set; } = 1; // 🟢 新增：預設每 1 天執行一次
        public static DateTime LastBackupDate { get; private set; } = DateTime.MinValue;

        public static void LoadConfig()
        {
            string dbPath = DataManager.GetSysSetting("BackupPath", "");
            string dbKeep = DataManager.GetSysSetting("BackupKeepCount", "");
            string dbInterval = DataManager.GetSysSetting("BackupIntervalDays", ""); // 🟢 讀取自訂的備份頻率
            string dbDate = DataManager.GetSysSetting("LastBackupDate", "");

            if (!string.IsNullOrEmpty(dbPath)) BackupPath = dbPath;
            if (int.TryParse(dbKeep, out int count)) KeepCount = count;
            if (int.TryParse(dbInterval, out int interval)) BackupIntervalDays = interval;
            if (DateTime.TryParse(dbDate, out DateTime date)) LastBackupDate = date;

            if (!Directory.Exists(BackupPath)) Directory.CreateDirectory(BackupPath);
        }

        // 🟢 修改：加入 intervalDays 參數
        public static void SaveConfig(string path, int count, int intervalDays)
        {
            BackupPath = path;
            KeepCount = count;
            BackupIntervalDays = intervalDays;
            
            DataManager.SetSysSetting("BackupPath", path);
            DataManager.SetSysSetting("BackupKeepCount", count.ToString());
            DataManager.SetSysSetting("BackupIntervalDays", intervalDays.ToString());
            
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
            
            // 🟢 修改：以使用者自訂的 BackupIntervalDays 作為判斷基準
            if ((DateTime.Today - LastBackupDate).TotalDays >= BackupIntervalDays)
            {
                // 系統性優化 5：備份併發控制 (避免 5 人同時啟動造成 I/O 風暴)
                string lockTimeStr = DataManager.GetSysSetting("BackupLockTime", "");
                
                // 檢查是否有人正在備份 (設定鎖定時間在 10 分鐘內視為鎖定中)
                if (DateTime.TryParse(lockTimeStr, out DateTime lockTime))
                {
                    if ((DateTime.Now - lockTime).TotalMinutes < 10)
                    {
                        // 已經有其他使用者的電腦正在執行備份，本機自動跳過
                        return;
                    }
                }

                // 搶佔備份鎖 (寫入當前時間)
                DataManager.SetSysSetting("BackupLockTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                try
                {
                    ExecuteBackup();
                }
                finally
                {
                    // 備份完成後解除鎖定 (清空時間)
                    DataManager.SetSysSetting("BackupLockTime", "");
                }
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
                // 背景執行時跳出警告可能會稍微干擾操作，但對防護來說是必要的提示
                MessageBox.Show("資料庫備份失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
