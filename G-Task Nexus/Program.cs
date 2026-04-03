using System;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace GTaskNexus
{
    internal static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private const int SW_RESTORE = 9;
        private static Mutex appMutex = new Mutex(true, "Global\\GTaskNexus_SingleInstance_Mutex");

        [STAThread]
        static void Main()
        {
            // 檢查是否已在執行
            if (!appMutex.WaitOne(TimeSpan.Zero, true))
            {
                // 尋找已存在的視窗 (標題需與 MainForm 的 Text 一致)
                IntPtr hWnd = FindWindow(null, "G-Task Nexus | 雙向同步中心");
                if (hWnd != IntPtr.Zero)
                {
                    ShowWindow(hWnd, SW_RESTORE); // 如果最小化則還原
                    SetForegroundWindow(hWnd);    // 移至最前方
                }
                return; 
            }

            try
            {
                if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware();

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            finally
            {
                appMutex.ReleaseMutex();
            }
        }
    }
}
