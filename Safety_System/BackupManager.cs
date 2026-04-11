/// FILE: Safety_System/BackupManager.cs ///
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Safety_System
{
    public static class BackupManager
    {
        // 設定檔存放路徑
        private static readonly string ConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "backup_config.txt");
        
        // 預設值
        public static string BackupPath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB_Backups");
        public static int KeepCount { get; private set; } = 4; // 預設保留 4 份 (約一個月)
        public static DateTime LastBackupDate { get; private set; } = DateTime.MinValue;

        public static void LoadConfig()
        {
            if (File.Exists(ConfigFile))
            {
                try {
                    string[] lines = File.ReadAllLines(ConfigFile, Encoding.UTF8);
                    if (lines.Length >= 1 && !string.IsNullOrWhiteSpace(lines[0])) BackupPath = lines[0];
                    if (lines.Length >= 2 && int.TryParse(lines[1], out int count)) KeepCount = count;
                    if (lines.Length >= 3 && DateTime.TryParse(lines[2], out DateTime date)) LastBackupDate = date;
                } catch { }
            }
            if (!Directory.Exists(BackupPath)) Directory.CreateDirectory(BackupPath);
        }

        public static void SaveConfig(string path, int count)
        {
            BackupPath = path;
            KeepCount = count;
            UpdateConfigDate(LastBackupDate);
            if (!Directory.Exists(BackupPath)) Directory.CreateDirectory(BackupPath);
        }

        private static void UpdateConfigDate(DateTime date)
        {
            LastBackupDate = date;
            File.WriteAllLines(ConfigFile, new[] { BackupPath, KeepCount.ToString(), LastBackupDate.ToString("yyyy-MM-dd") }, Encoding.UTF8);
        }

        // 🟢 給 MainForm 呼叫：程式啟動時自動檢查
        public static void RunAutoBackup()
        {
            LoadConfig();

            // 判斷是否超過 7 天
            if ((DateTime.Today - LastBackupDate).TotalDays >= 7)
            {
                ExecuteBackup();
            }
        }

        // 🟢 執行備份的核心邏輯 (包含建立資料夾與複製)
        public static void ExecuteBackup()
        {
            try
            {
                if (!Directory.Exists(DataManager.BasePath)) return;

                // 1. 建立當天日期的資料夾 (例如: 20231027_0930)
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

        // 🟢 清理超過保留份數的舊資料夾
        private static void CleanupOldBackups()
        {
            try
            {
                var directories = new DirectoryInfo(BackupPath).GetDirectories()
                                    .OrderByDescending(d => d.CreationTime)
                                    .ToList();

                // 如果資料夾數量超過保留數量，刪除較舊的
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
