using System;
using System.IO;
using System.Text;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly object _fileLock = new object();
        private const string InspectionFile = "InspectionData.txt";

        public static void SaveInspectionRecord(string date, string location, string inspector, string status)
        {
            // 欄位以 | 分隔
            string record = string.Format("{0}|{1}|{2}|{3}{4}", date, location, inspector, status, Environment.NewLine);
            
            lock (_fileLock)
            {
                // 使用 UTF8 編碼確保中文不掉碼
                File.AppendAllText(InspectionFile, record, Encoding.UTF8);
            }
        }
    }
}
