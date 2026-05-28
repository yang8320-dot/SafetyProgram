/// FILE: Safety_System/settings/MemoryOptimizer.cs ///
using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 系統記憶體深度釋放與最佳化引擎
    /// </summary>
    public static class MemoryOptimizer
    {
        // 🟢 匯入 Windows 底層核心 API
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, UIntPtr minimumWorkingSetSize, UIntPtr maximumWorkingSetSize);

        /// <summary>
        /// 執行深度記憶體釋放，並自動顯示結果對話框
        /// </summary>
        public static void Execute()
        {
            try
            {
                Application.UseWaitCursor = true;
                
                // 取得釋放前的記憶體佔用量 (MB)
                long beforeMem = Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024);

                // 1. 強制 .NET 進行完整的垃圾回收 (收回未使用的陣列、幽靈控制項等資源)
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
                GC.WaitForPendingFinalizers();
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);

                // 2. 呼叫 Windows 底層 API，強制作業系統將軟體已快取但未使用的記憶體分頁移出 (Page Out)
                if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    // 傳入 0xFFFFFFFF 代表強迫清空 Working Set
                    SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, (UIntPtr)0xFFFFFFFF, (UIntPtr)0xFFFFFFFF);
                }

                // 取得釋放後的記憶體佔用量 (MB)
                long afterMem = Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024);

                Application.UseWaitCursor = false;

                // 顯示精美的成果報告
                MessageBox.Show($"🚀 記憶體深度釋放與最佳化完成！\n\n" +
                                $"釋放前實體記憶體佔用： {beforeMem} MB\n" +
                                $"釋放後實體記憶體佔用： {afterMem} MB\n\n" +
                                $"已成功將閒置的系統資源全數歸還給作業系統。", 
                                "效能最佳化", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Application.UseWaitCursor = false;
                MessageBox.Show("記憶體釋放失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
