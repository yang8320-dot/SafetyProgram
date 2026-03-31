using System;
using System.IO;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly object _fileLock = new object();
        private const string InspectionFile = "InspectionData.txt";

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            // 格式化資料，以 | 分隔
            string record = $"{date}|{location}|{inspector}|{status}{Environment.NewLine}";

            lock (_fileLock)
            {
                File.AppendAllText(InspectionFile, record);
            }
        }
    }
}
