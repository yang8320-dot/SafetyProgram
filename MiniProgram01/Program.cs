/*
 * 檔案功能：程式進入點 (Entry Point)，負責啟動主視窗、確保單一執行個體 (Mutex) 與設定 DPI 縮放支援
 * 對應選單名稱：無 (系統啟動核心)
 * 對應資料庫名稱：無
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
        // 嘗試取得 Mutex 的所有權
        // TimeSpan.Zero 表示不等待，如果已經有其他實體拿走識別證了，就立刻回傳 false
        if (mutex.WaitOne(TimeSpan.Zero, true))
        {
            try
            {
                // 只有拿到識別證 (第一次開啟) 的程式才能執行這裡
                if (Environment.OSVersion.Version.Major >= 6)
                {
                    SetProcessDPIAware(); // 啟用 DPI 縮放支援
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                // 啟動我們設計好的 iOS 風格主視窗
                Application.Run(new MainForm());
            }
            finally
            {
                // 程式關閉時釋放識別證
                mutex.ReleaseMutex();
            }
        }
        else
        {
            // 如果沒拿到識別證，代表程式【已經在執行中】了
            // 這裡直接退出，不做任何事，避免產生第二個系統列常駐圖示
            MessageBox.Show("整合通知中心已在背景執行中！\n請查看桌面右下角系統列 (Tray) 的常駐圖示，或按下熱鍵 (Ctrl+1) 喚醒。", 
                            "提示", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Information);
            return;
        }
    }
}
