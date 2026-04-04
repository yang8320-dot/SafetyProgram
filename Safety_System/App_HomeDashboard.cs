using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 首頁綜合數據看板模組
    /// 對應選單：頁首
    /// 功能：未來作為跨模組數據 (COUNT/SUM) 的聯集展示區
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

            // 大標題 (遵守無 Unicode 表情符號規範，使用 ASCII 符號)
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
                Text = "歡迎使用工安系統看板！此頁面保留為綜合數據展示區，未來可整合各資料庫統計圖表與摘要。",
                Font = new Font("Microsoft JhengHei UI", 14F),
                AutoSize = true,
                Location = new Point(25, 70)
            };

            // 預留區塊：跨模組數據摘要
            GroupBox boxSummary = new GroupBox
            {
                Text = "系統數據總覽 (開發保留區)",
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Location = new Point(20, 120),
                Size = new Size(800, 300),
                Padding = new Padding(15) // 框內與文字間隔 15
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
            mainPanel.Controls.Add(boxSummary);

            return mainPanel;
        }
    }
}
