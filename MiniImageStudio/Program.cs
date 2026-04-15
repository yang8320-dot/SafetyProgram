/* * 功能：程式進入點 (加入 DPI 支援解決模糊)
 */
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace MiniImageStudio {
    static class Program {
        // 導入系統 API 以解決字型模糊
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main() {
            if (Environment.OSVersion.Version.Major >= 6) {
                SetProcessDPIAware();
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
