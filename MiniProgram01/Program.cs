using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

public static class Program {
    // 匯入 Windows API 以支援高 DPI 顯示，避免介面放大後模糊
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDPIAware();

    [DllImport("shcore.dll")]
    private static extern int SetProcessDpiAwareness(int awareness);

    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    [STAThread] 
    public static void Main() { 
        if (mutex.WaitOne(TimeSpan.Zero, true)) {
            try {
                // 啟用 Windows 8.1+ 的進階 DPI 支援
                if (Environment.OSVersion.Version.Major >= 6) {
                    try {
                        SetProcessDpiAwareness(1); // System DPI aware
                    } catch {
                        SetProcessDPIAware();
                    }
                }
                
                Application.EnableVisualStyles(); 
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm()); 
            } finally {
                mutex.ReleaseMutex();
            }
        } else {
            // 程式已在執行中，直接退出
            return;
        }
    }
}
