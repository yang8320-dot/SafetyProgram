/* * 功能：程式進入點 (加入 DPI 支援與防止重複開啟機制)
 */
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

namespace MiniImageStudio {
    static class Program {
        // 導入系統 API 以解決字型模糊
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        // 導入系統 API 以將視窗帶到最上層
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // 新增 ShowWindow API 來強制還原被最小化的視窗
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        // 宣告 Mutex 防止重複執行 (使用唯一的 GUID)
        static Mutex mutex = new Mutex(true, "{MiniImageStudio-Pro-Unique-Instance}");

        [STAThread]
        static void Main() {
            // 檢查是否已經有同名的 Mutex 存在
            if (mutex.WaitOne(TimeSpan.Zero, true)) {
                if (Environment.OSVersion.Version.Major >= 6) {
                    SetProcessDPIAware();
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
                mutex.ReleaseMutex();
            } else {
                // 如果程式已執行，找到舊視窗並把它移到最上層
                // 注意：這裡的標題必須與 MainForm.Text 完全一致
                IntPtr hwnd = FindWindow(null, "圖片工具程式 - 專業終極版");
                if (hwnd != IntPtr.Zero) {
                    ShowWindow(hwnd, SW_RESTORE); // 強制還原最小化的視窗
                    SetForegroundWindow(hwnd);    // 帶到最上層
                }
            }
        }
    }
}
