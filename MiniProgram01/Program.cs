using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

public static class Program {
    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    // 宣告一個全域的 Mutex，使用一個獨一無二的 GUID 名稱來作為這個程式的專屬識別證
    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    [STAThread] 
    public static void Main() { 
        // 嘗試取得 Mutex 的所有權
        // TimeSpan.Zero 表示不等待，如果已經有其他實體拿走識別證了，就立刻回傳 false
        if (mutex.WaitOne(TimeSpan.Zero, true)) {
            try {
                // 只有拿到識別證 (第一次開啟) 的程式才能執行這裡
                if (Environment.OSVersion.Version.Major >= 6) { SetProcessDPIAware(); }
                Application.EnableVisualStyles(); 
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm()); 
            } finally {
                // 程式關閉時釋放識別證
                mutex.ReleaseMutex();
            }
        } else {
            // 如果沒拿到識別證，代表程式【已經在執行中】了
            // 這裡直接退出，不做任何事 (就不會產生第二個常駐圖示)
            return;
        }
    }
}
