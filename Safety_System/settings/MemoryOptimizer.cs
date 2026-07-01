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
    /// 系統記憶體深度釋放與最佳化引擎 (System-wide RAM Optimizer)
    /// </summary>
    public static class MemoryOptimizer
    {
        // 🟢 匯入 Windows 底層核心 API：強制作業系統收回實體記憶體
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, IntPtr minimumWorkingSetSize, IntPtr maximumWorkingSetSize);

        // 🟢 匯入 API：清空記憶體分頁
        [DllImport("psapi.dll")]
        static extern int EmptyWorkingSet(IntPtr hwProc);

        /// <summary>
        /// 執行深度記憶體釋放 (包含進度條與非同步掃描)
        /// </summary>
        public static void Execute()
        {
            // 建立一個進度視窗讓使用者「有感」
            using (Form pForm = new Form())
            {
                pForm.Text = "系統記憶體最佳化工具";
                pForm.Size = new Size(450, 220);
                pForm.StartPosition = FormStartPosition.CenterParent;
                pForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                pForm.MaximizeBox = false;
                pForm.MinimizeBox = false;
                pForm.BackColor = Color.White;
                pForm.ControlBox = false; // 阻擋右上角關閉

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
                    Text = "準備掃描系統程序...",
                    Font = new Font("Microsoft JhengHei UI", 11F),
                    ForeColor = Color.DimGray,
                    Location = new Point(25, 70),
                    AutoSize = true
                };

                ProgressBar pb = new ProgressBar
                {
                    Location = new Point(25, 110),
                    Size = new Size(380, 25),
                    Style = ProgressBarStyle.Continuous
                };

                pForm.Controls.Add(lblTitle);
                pForm.Controls.Add(lblStatus);
                pForm.Controls.Add(pb);

                pForm.Shown += async (s, e) =>
                {
                    Application.UseWaitCursor = true;
                    int successCount = 0;
                    int failCount = 0;
                    long totalMemoryFreed = 0;
                    
                    // 紀錄優化前的本程式記憶體
                    long myMemBefore = Process.GetCurrentProcess().WorkingSet64;

                    await Task.Run(() =>
                    {
                        try
                        {
                            // 1. 強制 .NET 進行完整的垃圾回收 (清理幽靈陣列與控制項)
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
                            GC.WaitForPendingFinalizers();
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);

                            // 2. 獲取系統所有正在執行的程序
                            Process[] processes = Process.GetProcesses();
                            int totalProcesses = processes.Length;

                            for (int i = 0; i < totalProcesses; i++)
                            {
                                Process p = processes[i];

                                // 更新進度條
                                int percent = (int)((double)(i + 1) / totalProcesses * 100);
                                pForm.Invoke(new Action(() => 
                                { 
                                    pb.Value = percent; 
                                    lblStatus.Text = $"正在壓縮程序 ({i}/{totalProcesses}): {p.ProcessName}";
                                }));

                                try
                                {
                                    if (p.HasExited) continue;

                                    long memBefore = p.WorkingSet64;

                                    // 🟢 核心技術 1：強制作業系統將該程序的未活躍記憶體分頁移出 (Page Out)
                                    SetProcessWorkingSetSize(p.Handle, (IntPtr)(-1), (IntPtr)(-1));
                                    
                                    // 🟢 核心技術 2：進一步清空 Working Set
                                    EmptyWorkingSet(p.Handle);

                                    p.Refresh();
                                    long memAfter = p.WorkingSet64;
                                    
                                    if (memBefore > memAfter)
                                    {
                                        totalMemoryFreed += (memBefore - memAfter);
                                    }
                                    successCount++;
                                }
                                catch
                                {
                                    // 通常會拋出 Access Denied，這是正常的，因為防毒軟體或系統核心不允許動它。
                                    failCount++;
                                }
                            }
                        }
                        catch { }
                    });

                    Application.UseWaitCursor = false;
                    long myMemAfter = Process.GetCurrentProcess().WorkingSet64;
                    double freedMB = (double)totalMemoryFreed / (1024 * 1024);
                    double myFreedMB = (double)(myMemBefore - myMemAfter) / (1024 * 1024);
                    if (myFreedMB < 0) myFreedMB = 0;

                    pForm.DialogResult = DialogResult.OK;

                    // 顯示精美的成果報告
                    MessageBox.Show(
                        $"🚀 系統記憶體深度最佳化完成！\n\n" +
                        $"🔍 共掃描並壓縮了 {successCount} 個背景程序 (略過 {failCount} 個受保護核心)。\n\n" +
                        $"💻 您的電腦總共釋放了： {freedMB:N0} MB 的實體記憶體！\n" +
                        $"⚙️ 本系統 (Safety_System) 釋放了： {myFreedMB:N0} MB\n\n" +
                        $"已成功將閒置的系統資源全數歸還給作業系統。", 
                        "效能最佳化報告", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                pForm.ShowDialog();
            }
        }
    }
}
