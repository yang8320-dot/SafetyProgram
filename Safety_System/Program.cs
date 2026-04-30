/// FILE: Safety_System/Program.cs ///
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
            bool createdNew;
            using (Mutex mutex = new Mutex(true, "SafetySystem_Unique_Mutex_Name", out createdNew))
            {
                if (createdNew)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    // 🟢 在顯示主畫面之前，先進行軟體啟用與權限認證
                    if (!LicenseManager.VerifyLicense())
                    {
                        MessageBox.Show("未申請使用此軟體功能！請洽詢軟體管理者。", "權限不足或授權到期", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return; // 驗證失敗，直接終止程式
                    }

                    Application.Run(new MainForm());
                }
                else
                {
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
    }
}
