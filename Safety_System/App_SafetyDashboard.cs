using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyDashboard
    {
        public Control GetView()
        {
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, Padding = new Padding(30) };
            
            Label lblTitle = new Label { 
                Text = "🛡️ 工安管理儀表版", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                AutoSize = true, Location = new Point(30, 30) 
            };

            // 範例：零災害天數看板
            Panel pnlCard = new Panel { 
                Size = new Size(400, 150), Location = new Point(35, 100), 
                BackColor = Color.FromArgb(45, 45, 45), // 深色質感
            };
            Label lblDaysTitle = new Label { Text = "安全生產天數", ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F), Location = new Point(20, 20), AutoSize = true };
            Label lblDays = new Label { Text = "365 天", ForeColor = Color.LimeGreen, Font = new Font("Impact", 40F), Location = new Point(20, 60), AutoSize = true };
            
            pnlCard.Controls.AddRange(new Control[] { lblDaysTitle, lblDays });
            pnl.Controls.AddRange(new Control[] { lblTitle, pnlCard });
            
            return pnl;
        }
    }
}
