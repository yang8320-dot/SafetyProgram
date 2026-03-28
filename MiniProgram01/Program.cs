using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public static class Program {
    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    [STAThread] 
    public static void Main() { 
        if (Environment.OSVersion.Version.Major >= 6) { SetProcessDPIAware(); }
        Application.EnableVisualStyles(); 
        Application.Run(new MainForm()); 
    }
}
