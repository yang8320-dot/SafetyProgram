/* * 功能：存取操作歷史紀錄
 * 對應選單名稱：系統紀錄
 * 對應資料庫名稱：HistoryDB (純文字儲存)
 * 對應資料表名稱：history.txt
 */
using System;
using System.IO;

namespace MiniImageStudio {
    public static class App_History {
        // 設定紀錄檔存放的名稱
        private static string filePath = "history.txt";

        public static void WriteLog(string action) {
            try {
                // 強制日期格式 YYYY-MM-DD HH:mm:ss
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string logEntry = $"{timestamp}|{action}{Environment.NewLine}";
                
                // 寫入純文字檔
                File.AppendAllText(filePath, logEntry);
            }
            catch { 
                // 如果遇到權限問題或檔案被佔用，直接略過，避免主程式崩潰
            }
        }
    }
}
