using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    static class Program
    {
        // 匯入 Windows API 用於喚醒視窗
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9; // 強制還原並顯示視窗的指令

        [STAThread]
        static void Main()
        {
            bool createdNew;
            // 使用 Mutex 確保系統中只有一個此名稱的程式在執行
            using (Mutex mutex = new Mutex(true, "SafetySystem_Unique_Mutex_Name", out createdNew))
            {
                if (createdNew)
                {
                    // 第一次執行，正常啟動
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new MainForm());
                }
                else
                {
                    // 程式已在執行，尋找原本的視窗
                    Process current = Process.GetCurrentProcess();
                    foreach (Process process in Process.GetProcessesByName(current.ProcessName))
                    {
                        if (process.Id != current.Id)
                        {
                            IntPtr handle = process.MainWindowHandle;
                            if (handle != IntPtr.Zero)
                            {
                                // 1. 還原視窗（避免視窗處於最小化狀態）
                                ShowWindow(handle, SW_RESTORE);
                                // 2. 將視窗帶到最前端
                                SetForegroundWindow(handle);
                            }
                            break;
                        }
                    }
                    // 執行完畢，結束當前這個重複的實例，不跳提示視窗
                }
            }
        }
    }
}
