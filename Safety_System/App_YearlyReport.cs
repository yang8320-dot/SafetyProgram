using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_YearlyReport
    {
        public Control GetView()
        {
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.GhostWhite };
            
            Label lblTitle = new Label
            {
                Text = "數據分析 - 年報表中心",
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold),
                Location = new Point(30, 30),
                AutoSize = true
            };
            pnl.Controls.Add(lblTitle);

            Label lblInfo = new Label
            {
                Text = "本頁面對應 App_YearlyReport.cs, 資料尚在建立中\n(未來功能：年度工安趨勢分析、各廠區巡檢數據對比)",
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Italic),
                Location = new Point(30, 100),
                AutoSize = true,
                ForeColor = Color.DimGray
            };
            pnl.Controls.Add(lblInfo);

            return pnl;
        }
    }
}
