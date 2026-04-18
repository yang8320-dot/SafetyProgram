using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

public static class Program {
    [span_2](start_span)// 匯入 Windows API 以支援高 DPI 顯示，避免介面放大後模糊[span_2](end_span)
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDPIAware();

    [DllImport("shcore.dll")]
    private static extern int SetProcessDpiAwareness(int awareness);

    [span_3](start_span)// 宣告一個全域的 Mutex，確保程式單一執行個體[span_3](end_span)
    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    [STAThread] 
    public static void Main() { 
        [span_4](start_span)// 嘗試取得 Mutex 的所有權[span_4](end_span)
        if (mutex.WaitOne(TimeSpan.Zero, true)) {
            try {
                [span_5](start_span)// 優先使用 Windows 8.1+ 的進階 DPI 支援[span_5](end_span)
                if (Environment.OSVersion.Version.Major >= 6) {
                    try {
                        SetProcessDpiAwareness(1); // Process_System_DPI_Aware
                    } catch {
                        SetProcessDPIAware();
                    }
                }
                
                Application.EnableVisualStyles(); 
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            } finally {
                [span_6](start_span)// 程式關閉時釋放識別證[span_6](end_span)
                mutex.ReleaseMutex();
            }
        } else {
            [span_7](start_span)// 如果程式已在執行中，直接退出[span_7](end_span)
            return;
        }
    }
}
