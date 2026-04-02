using System;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    static class Program
    {
        // 宣告一個全域的 Mutex
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            // 給你的程式一個唯一的識別碼
            const string appName = "SafetySystem_SingleInstance_Mutex_v1";
            bool createdNew;

            // 嘗試取得 Mutex，若 createdNew 為 false，代表已經有一個實體在執行
            mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                MessageBox.Show("工安系統已經在執行中，請勿重複啟動！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // 直接退出程式
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // 程式啟動時，先載入設定好的資料庫路徑
            DataManager.LoadConfig();

            Application.Run(new MainForm());
        }
    }
}
