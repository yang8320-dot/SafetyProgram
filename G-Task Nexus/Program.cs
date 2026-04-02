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

        // 宣告全域互斥鎖 (Mutex)，使用獨一無二的 GUID 確保不衝突
        private static Mutex appMutex = new Mutex(true, "Global\\GTaskNexus_SingleInstance_Mutex");

        [STAThread]
        static void Main()
        {
            // 判斷是否已經有相同的 Mutex 存在 (防止重複開啟)
            if (!appMutex.WaitOne(TimeSpan.Zero, true))
            {
                MessageBox.Show("G-Task Nexus 已經在背景執行中！\n請按下快捷鍵「Ctrl + 2」喚醒程式。", 
                                "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return; // 直接結束這次的開啟
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
                // 程式結束時釋放 Mutex
                appMutex.ReleaseMutex();
            }
        }
    }
}
