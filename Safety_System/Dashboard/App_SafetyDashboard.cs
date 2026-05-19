using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyDashboard
    {
        public Control GetView()
        {
            // 🟢 增加頂部 Padding 防止被遮擋
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            
            Label lblTitle = new Label { 
                Text = "🛡️ 工安管理儀表版", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                AutoSize = true, Location = new Point(30, 20) 
            };

            GroupBox box = new GroupBox { Text = "安全指標", Size = new Size(400, 200), Location = new Point(30, 80), Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblDays = new Label { Text = "零災害累計天數：365 天", ForeColor = Color.Green, AutoSize = true, Location = new Point(20, 40) };
            box.Controls.Add(lblDays);

            pnl.Controls.AddRange(new Control[] { lblTitle, box });
            main.Controls.Add(pnl);
            return main;
        }
    }
}
