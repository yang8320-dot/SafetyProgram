using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Drawing;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly object _fileLock = new object();
        
        // 系統預設資料存放路徑 (預設為 .exe 所在目錄)
        public static string BasePath { get; private set; } = AppDomain.CurrentDomain.BaseDirectory;
        private const string ConfigFile = "sys_config.txt";

        // 動態組合出完整的檔案路徑
        private static string GetFilePath(string fileName)
        {
            return Path.Combine(BasePath, fileName);
        }

        /// <summary>
        /// 載入系統設定 (讀取使用者指定的資料庫路徑)
        /// </summary>
        public static void LoadConfig()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigFile);
            if (File.Exists(configPath))
            {
                string savedPath = File.ReadAllText(configPath, Encoding.UTF8).Trim();
                if (Directory.Exists(savedPath))
                {
                    BasePath = savedPath; // 若路徑存在，則切換過去
                }
            }
        }

        /// <summary>
        /// 設定並記憶新的資料庫路徑
        /// </summary>
        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigFile);
            File.WriteAllText(configPath, newPath, Encoding.UTF8);
        }

        // --- 以下為各項資料的讀寫功能 (已全部改用 GetFilePath) ---

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            string record = string.Format("{0}|{1}|{2}|{3}{4}", date, location, inspector, status, Environment.NewLine);
            lock (_fileLock)
            {
                File.AppendAllText(GetFilePath("InspectionData.txt"), record, Encoding.UTF8);
            }
        }

        public static void SaveTodo(string task, string date, string colorHex)
        {
            string record = string.Format("{0}|{1}|{2}{3}", task, date, colorHex, Environment.NewLine);
            lock (_fileLock)
            {
                File.AppendAllText(GetFilePath("TodoList.txt"), record, Encoding.UTF8);
            }
        }

        public static List<string[]> LoadTodos()
        {
            List<string[]> list = new List<string[]>();
            lock (_fileLock)
            {
                string targetFile = GetFilePath("TodoList.txt");
                if (File.Exists(targetFile))
                {
                    string[] lines = File.ReadAllLines(targetFile, Encoding.UTF8);
                    foreach (string line in lines)
                    {
                        if (!string.IsNullOrEmpty(line))
                            list.Add(line.Split('|'));
                    }
                }
            }
            return list;
        }

        public static void OverwriteTodos(List<string[]> allData)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in allData)
            {
                sb.AppendLine(string.Format("{0}|{1}|{2}", item[0], item[1], item[2]));
            }
            lock (_fileLock)
            {
                File.WriteAllText(GetFilePath("TodoList.txt"), sb.ToString(), Encoding.UTF8);
            }
        }
    }
}
