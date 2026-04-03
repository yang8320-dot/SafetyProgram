using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_Report
    {
        /// <summary>
        /// 取得月報表畫面
        /// </summary>
        public Control GetMonthlyReportView()
        {
            return CreateBaseView("數據分析 - 月報表", Color.AliceBlue);
        }

        /// <summary>
        /// 取得年報表畫面
        /// </summary>
        public Control GetYearlyReportView()
        {
            return CreateBaseView("數據分析 - 年報表", Color.GhostWhite);
        }

        private Control CreateBaseView(string title, Color bgColor)
        {
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = bgColor };
            
            Label lblTitle = new Label
            {
                Text = title,
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold),
                Location = new Point(30, 30),
                AutoSize = true
            };
            pnl.Controls.Add(lblTitle);

            Label lblPlaceholder = new Label
            {
                Text = "報表數據彙整功能開發中...\n(此處未來將結合 DataManager 讀取 txt 並製作統計圖表)",
                Font = new Font("Microsoft JhengHei UI", 14F),
                Location = new Point(30, 100),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            pnl.Controls.Add(lblPlaceholder);

            return pnl;
        }
    }
}
