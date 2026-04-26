/// FILE: Safety_System/BackupManager.cs ///
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public static class BackupManager
    {
        // 預設值
        public static string BackupPath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB_Backups");
        public static int KeepCount { get; private set; } = 4; // 預設保留 4 份 (約一個月)
        public static DateTime LastBackupDate { get; private set; } = DateTime.MinValue;

        // 🟢 改用 DataManager 全新 DB 存取機制
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

            // 判斷是否超過 7 天
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

                // 1. 建立當天日期的資料夾
                string folderName = DateTime.Now.ToString("yyyyMMdd_HHmm");
                string targetDir = Path.Combine(BackupPath, folderName);
                Directory.CreateDirectory(targetDir);

                // 2. 複製所有 .sqlite 檔案
                string[] files = Directory.GetFiles(DataManager.BasePath, "*.sqlite");
                foreach (var file in files)
                {
                    string destFile = Path.Combine(targetDir, Path.GetFileName(file));
                    File.Copy(file, destFile, true);
                }

                // 🟢 將系統核心配置也一併備份
                string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                if (File.Exists(sysDbPath))
                {
                    File.Copy(sysDbPath, Path.Combine(targetDir, "SystemConfig.sqlite"), true);
                }

                // 3. 更新備份紀錄日期為今天
                UpdateConfigDate(DateTime.Today);

                // 4. 清理舊備份
                CleanupOldBackups();
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料庫備份失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
