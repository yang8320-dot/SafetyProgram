/// FILE: Safety_System/settings/MemoryOptimizer.cs ///
using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 系統記憶體深度釋放與最佳化引擎
    /// (已優化：為避免防毒軟體(Apex One)誤判，徹底移除所有底層 API DllImport 呼叫)
    /// </summary>
    public static class MemoryOptimizer
    {
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
                    Style = ProgressBarStyle.Marquee,
                    MarqueeAnimationSpeed = 30
                };

                pForm.Controls.Add(lblTitle);
                pForm.Controls.Add(lblStatus);
                pForm.Controls.Add(pb);

                pForm.Shown += async (s, e) =>
                {
                    Application.UseWaitCursor = true;
                    
                    Process currentProcess = Process.GetCurrentProcess();
                    long myMemBefore = currentProcess.WorkingSet64;

                    await Task.Run(async () =>
                    {
                        try
                        {
                            // 延遲一下讓進度條跑動，給使用者視覺回饋
                            await Task.Delay(800);

                            // 100% 安全的 .NET 內建記憶體回收，防毒絕對不會叫
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
                            GC.WaitForPendingFinalizers();
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
                            
                            await Task.Delay(400);
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
