using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9; 

        [STAThread]
        static void Main()
        {
            // 🟢 加入全域例外捕捉，防止程式靜默崩潰 (閃退)
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            try
            {
                bool createdNew;
                // 使用 Mutex 防止程式重複開啟
                using (Mutex mutex = new Mutex(true, "SafetySystem_Unique_Mutex_Name", out createdNew))
                {
                    if (createdNew)
                    {
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);

                        // 🟢 在顯示主畫面之前，先進行軟體啟用與權限認證
                        if (!LicenseManager.VerifyLicense())
                        {
                            MessageBox.Show("您的電腦帳號尚未在授權名單內！\n\n(若要測試，請至 LicenseManager.cs 暫時修改 VerifyLicense() 回傳 true)", 
                                            "權限不足或授權到期", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return; // 驗證失敗，直接終止程式
                        }

                        // 啟動主畫面
                        Application.Run(new MainForm());
                    }
                    else
                    {
                        // 程式已在執行中，嘗試喚醒舊視窗
                        Process current = Process.GetCurrentProcess();
                        foreach (Process process in Process.GetProcessesByName(current.ProcessName))
                        {
                            if (process.Id != current.Id)
                            {
                                IntPtr handle = process.MainWindowHandle;
                                if (handle != IntPtr.Zero)
                                {
                                    ShowWindow(handle, SW_RESTORE);
                                    SetForegroundWindow(handle);
                                }
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 捕捉到致命錯誤時顯示
                ShowFatalError(ex);
            }
        }

        // 🟢 處理 UI 執行緒的未捕捉錯誤
        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            ShowFatalError(e.Exception);
        }

        // 🟢 處理非 UI 執行緒的未捕捉錯誤
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                ShowFatalError(ex);
            }
        }

        // 🟢 統一的錯誤訊息顯示方法
        private static void ShowFatalError(Exception ex)
        {
            string errorMsg = $"系統啟動或執行時發生嚴重錯誤！\n\n" +
                              $"【錯誤訊息】：\n{ex.Message}\n\n" +
                              $"【錯誤追蹤】：\n{ex.StackTrace}";
            
            MessageBox.Show(errorMsg, "系統崩潰報告", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
