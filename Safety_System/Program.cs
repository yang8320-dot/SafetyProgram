/// FILE: Safety_System/Program.cs ///
using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    static class Program
    {
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

                        // 🟢 【關鍵修復】：在驗證權限、操作資料庫之前，先強制初始化與建立資料夾！
                        DataManager.LoadConfig();

                        // 🟢 在顯示主畫面之前，先進行軟體啟用與權限認證
                        if (!LicenseManager.VerifyLicense())
                        {
                            MessageBox.Show("您的電腦帳號尚未在授權名單內！\n\n(請聯絡系統管理員進行授權)", 
                                            "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return; // 驗證失敗，直接終止程式
                        }

                        // 啟動主畫面
                        Application.Run(new MainForm());
                    }
                    else
                    {
                        // 🟢 已拔除容易被誤判為惡意程式的 user32.dll (ShowWindow / SetForegroundWindow) 
                        // 改為溫和的提示視窗
                        MessageBox.Show("Safety System 已經在執行中了！\n請檢查您的 Windows 工具列或背景視窗。", 
                                        "系統已啟動", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            string innerMsg = ex.InnerException != null ? $"\n\n【內部詳細錯誤 (真正原因)】：\n{ex.InnerException.Message}" : "";
            
            string errorMsg = $"系統啟動或執行時發生嚴重錯誤！\n\n" +
                              $"【錯誤訊息】：\n{ex.Message}{innerMsg}\n\n" +
                              $"【錯誤追蹤】：\n{ex.StackTrace}";
            
            MessageBox.Show(errorMsg, "系統崩潰報告", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
