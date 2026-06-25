/// FILE: Safety_System/settings/MemoryOptimizer.cs ///
using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 系統記憶體深度釋放與最佳化引擎
    /// (已修正：為避免防毒軟體誤判為惡意程式，改為僅針對「本程式自身」進行記憶體壓縮與優化)
    /// </summary>
    public static class MemoryOptimizer
    {
        // 🟢 匯入 Windows 底層核心 API：強制作業系統收回實體記憶體
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, IntPtr minimumWorkingSetSize, IntPtr maximumWorkingSetSize);

        /// <summary>
        /// 執行深度記憶體釋放 (針對自身 Process)
        /// </summary>
        public static void Execute()
        {
            using (Form pForm = new Form())
            {
                pForm.Text = "系統記憶體最佳化工具";
                pForm.Size = new Size(450, 220);
                pForm.StartPosition = FormStartPosition.CenterParent;
                pForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                pForm.MaximizeBox = false;
                pForm.MinimizeBox = false;
                pForm.BackColor = Color.White;
                pForm.ControlBox = false; 

                Label lblTitle = new Label
                {
                    Text = "🚀 正在執行深度記憶體釋放...",
                    Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold),
                    ForeColor = Color.SteelBlue,
                    Location = new Point(20, 20),
                    AutoSize = true
                };

                Label lblStatus = new Label
                {
                    Text = "正在壓縮本系統資源...",
                    Font = new Font("Microsoft JhengHei UI", 11F),
                    ForeColor = Color.DimGray,
                    Location = new Point(25, 70),
                    AutoSize = true
                };

                ProgressBar pb = new ProgressBar
                {
                    Location = new Point(25, 110),
                    Size = new Size(380, 25),
                    Style = ProgressBarStyle.Marquee, // 改為跑馬燈模式
                    MarqueeAnimationSpeed = 30
                };

                pForm.Controls.Add(lblTitle);
                pForm.Controls.Add(lblStatus);
                pForm.Controls.Add(pb);

                pForm.Shown += async (s, e) =>
                {
                    Application.UseWaitCursor = true;
                    
                    // 紀錄優化前的本程式記憶體
                    Process currentProcess = Process.GetCurrentProcess();
                    long myMemBefore = currentProcess.WorkingSet64;

                    await Task.Run(() =>
                    {
                        try
                        {
                            // 1. 強制 .NET 進行完整的垃圾回收 (清理幽靈陣列與控制項)
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
                            GC.WaitForPendingFinalizers();
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);

                            // 2. 🟢 安全優化：只針對「本程式」進行 Windows 記憶體分頁壓縮，絕不碰觸其他程序
                            SetProcessWorkingSetSize(currentProcess.Handle, (IntPtr)(-1), (IntPtr)(-1));
                        }
                        catch { }
                    });

                    Application.UseWaitCursor = false;
                    currentProcess.Refresh();
                    long myMemAfter = currentProcess.WorkingSet64;
                    
                    double myFreedMB = (double)(myMemBefore - myMemAfter) / (1024 * 1024);
                    if (myFreedMB < 0) myFreedMB = 0;

                    pForm.DialogResult = DialogResult.OK;

                    MessageBox.Show(
                        $"🚀 系統記憶體最佳化完成！\n\n" +
                        $"⚙️ 本系統 (Safety_System) 成功釋放了： {myFreedMB:N1} MB 的記憶體！\n\n" +
                        $"已成功將閒置的資源歸還給作業系統，確保系統流暢運行。", 
                        "效能最佳化報告", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                pForm.ShowDialog();
            }
        }
    }
}
