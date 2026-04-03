using System;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    static class Program
    {
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            try 
            {
                // 1. 防止重複啟動檢查
                const string appName = "SafetySystem_SingleInstance_Mutex_v1";
                bool createdNew;
                mutex = new Mutex(true, appName, out createdNew);

                if (!createdNew)
                {
                    MessageBox.Show("工安系統已經在執行中！", "提示");
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                // 2. 初始化資料庫設定 (這步最容易報錯，所以放進 try 裡面)
                DataManager.LoadConfig();

                // 3. 啟動主視窗
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                // 如果程式啟動失敗，彈出具體錯誤原因
                MessageBox.Show("程式啟動發生錯誤：\n" + ex.Message + "\n\n堆疊追蹤：\n" + ex.StackTrace, "啟動失敗");
            }
            finally
            {
                if (mutex != null)
                {
                    mutex.ReleaseMutex();
                    mutex.Close();
                }
            }
        }
    }
}
