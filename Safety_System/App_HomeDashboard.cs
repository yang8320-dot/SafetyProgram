using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 首頁綜合數據看板模組
    /// 對應選單：頁首
    /// 功能：提供快捷導覽按鈕，以及未來作為跨模組數據 (COUNT/SUM) 的聯集展示區
    /// </summary>
    public class App_HomeDashboard
    {
        public Control GetView()
        {
            // 主容器：採用 Panel 配合自動滾動，提供更大的排版自由度
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                AutoScroll = true
            };

            // 大標題 
            Label lblTitle = new Label
            {
                Text = "[ 綜合數據看板 - 系統首頁 ]",
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(20, 20)
            };

            // 頁面說明
            Label lblDesc = new Label
            {
                Text = "歡迎使用工安系統看板！您可以從上方選單進入各模組，或使用下方快捷按鈕。",
                Font = new Font("Microsoft JhengHei UI", 14F),
                AutoSize = true,
                Location = new Point(25, 70)
            };

            // ==========================================
            // 🟢 新增區塊：快速導覽區 (擺放各大模組的捷徑按鈕)
            // ==========================================
            GroupBox boxQuickAccess = new GroupBox
            {
                Text = "快速導覽",
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Location = new Point(20, 120),
                Size = new Size(800, 150),
                Padding = new Padding(15)
            };

            // 使用 FlowLayoutPanel 讓裡面的大按鈕可以自動排列
            FlowLayoutPanel flpButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                WrapContents = true
            };

            // 建立「檢測數據」大按鈕
            Button btnTestDashboard = new Button 
            { 
                Text = "🔬 檢測數據", 
                Size = new Size(180, 80), 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold),
                BackColor = Color.SteelBlue, 
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Margin = new Padding(10)
            };

            // 點擊事件：尋找上層的 MainForm 並呼叫 LoadModule 切換畫面
            btnTestDashboard.Click += (s, e) => {
                Form parentForm = btnTestDashboard.FindForm();
                if (parentForm is MainForm mainForm) {
                    mainForm.LoadModule(new App_TestDashboard().GetView());
                }
            };

            // 將按鈕加入快速導覽區
            flpButtons.Controls.Add(btnTestDashboard);
            boxQuickAccess.Controls.Add(flpButtons);


            // ==========================================
            // 原有的預留區塊：跨模組數據摘要 (將 Y 座標往下移到 290)
            // ==========================================
            GroupBox boxSummary = new GroupBox
            {
                Text = "系統數據總覽 (開發保留區)",
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Location = new Point(20, 290), 
                Size = new Size(800, 300),
                Padding = new Padding(15) 
            };

            Label lblPlaceholder = new Label
            {
                Text = "載入中...\n\n(未來開發計畫：)\n1. 統計未結案的巡檢與異常事件\n2. 顯示本月各部門用水/空污申報進度\n3. 彙整近期需處理之消防設備巡檢任務",
                Font = new Font("Microsoft JhengHei UI", 12F),
                AutoSize = true,
                Location = new Point(20, 40)
            };

            boxSummary.Controls.Add(lblPlaceholder);

            // 將所有控制項加入主容器
            mainPanel.Controls.Add(lblTitle);
            mainPanel.Controls.Add(lblDesc);
            mainPanel.Controls.Add(boxQuickAccess); // 🟢 加入快捷區塊
            mainPanel.Controls.Add(boxSummary);

            return mainPanel;
        }
    }
}
