using System;
using System.IO;
using System.Text;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly object _fileLock = new object();
        private const string ConfigFile = "sys_config.txt";
        
        // 預設路徑為程式執行檔路徑
        public static string BasePath { get; private set; } = AppDomain.CurrentDomain.BaseDirectory;

        // 初始化：由 Program.cs 呼叫
        public static void LoadConfig()
        {
            if (File.Exists(ConfigFile))
            {
                string savedPath = File.ReadAllText(ConfigFile, Encoding.UTF8).Trim();
                if (Directory.Exists(savedPath)) BasePath = savedPath;
            }
        }

        // 更新路徑並寫入設定檔
        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            File.WriteAllText(ConfigFile, newPath, Encoding.UTF8);
        }

        // 取得完整檔案路徑的方法
        private static string GetFullPath(string fileName)
        {
            return Path.Combine(BasePath, fileName);
        }

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            string record = string.Format("{0}|{1}|{2}|{3}{4}", date, location, inspector, status, Environment.NewLine);
            lock (_fileLock)
            {
                File.AppendAllText(GetFullPath("InspectionData.txt"), record, Encoding.UTF8);
            }
        }
    }
}
