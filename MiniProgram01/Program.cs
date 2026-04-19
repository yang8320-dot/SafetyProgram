/*
 * 檔案功能：程式進入點 (加入全局錯誤捕捉網)
 */

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

public static class Program
{
    [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
    private static extern bool SetProcessDPIAware();

    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    [STAThread]
    public static void Main()
    {
        // 加入全局錯誤捕捉，防止程式默默閃退
        Application.ThreadException += (s, e) => 
            MessageBox.Show($"UI執行緒發生錯誤:\n{e.Exception.Message}\n\n{e.Exception.StackTrace}", "嚴重錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        
        AppDomain.CurrentDomain.UnhandledException += (s, e) => 
            MessageBox.Show($"發生非預期錯誤:\n{((Exception)e.ExceptionObject).Message}\n\n{((Exception)e.ExceptionObject).StackTrace}", "嚴重錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);

        try
        {
            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                if (Environment.OSVersion.Version.Major >= 6)
                {
                    SetProcessDPIAware();
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                // 這裡如果出錯，就會被下方的 catch 捕捉到
                DatabaseManager.InitializeDatabase();
                
                Application.Run(new MainForm());
            }
            else
            {
                MessageBox.Show("整合通知中心已在背景執行中！\n請按下 Ctrl+1 喚醒。", 
                                "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        catch (Exception ex)
        {
            // 捕捉初始化期間的任何崩潰
            MessageBox.Show($"程式啟動失敗！\n錯誤訊息: {ex.Message}\n\n詳細追蹤:\n{ex.StackTrace}", 
                            "啟動崩潰", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            try { mutex.ReleaseMutex(); } catch { }
        }
    }
}
