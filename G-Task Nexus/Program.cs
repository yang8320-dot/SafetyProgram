using System;
using System.Windows.Forms;

namespace GTaskNexus
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // 啟動主程式
            Application.Run(new MainForm());
        }
    }
}
