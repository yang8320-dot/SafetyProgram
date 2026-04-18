/*
 * 檔案功能：程式進入點 (Entry Point)，負責啟動主視窗、確保單一執行個體 (Mutex) 與設定 DPI 縮放支援
 * 對應選單名稱：無 (系統啟動核心)
 * 對應資料庫名稱：MainDB.sqlite (於此處觸發初始化)
 * 資料表名稱：無
 */

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

public static class Program
{
    // --- 引入 Windows API 確保高 DPI 螢幕下字體不模糊 (iOS 風格介面必備) ---
    [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
    private static extern bool SetProcessDPIAware();

    // --- 宣告全域 Mutex，確保程式為單一執行檔 (不會重複開啟) ---
    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    [STAThread]
    public static void Main()
    {
        if (mutex.WaitOne(TimeSpan.Zero, true))
        {
            try
            {
                if (Environment.OSVersion.Version.Major >= 6)
                {
                    SetProcessDPIAware(); // 啟用 DPI 縮放支援
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                // 【新增】：在介面啟動前，初始化 SQLite 資料庫與所有資料表
                DatabaseManager.InitializeDatabase();
                
                // 啟動我們設計好的 iOS 風格主視窗
                Application.Run(new MainForm());
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }
        else
        {
            MessageBox.Show("整合通知中心已在背景執行中！\n請查看桌面右下角系統列 (Tray) 的常駐圖示，或按下熱鍵 (Ctrl+1) 喚醒。", 
                            "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
    }
}
