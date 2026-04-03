using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterDashboard
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            
            Label lblTitle = new Label { 
                Text = "💧 水資源管理儀表版", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                AutoSize = true, Location = new Point(30, 20) 
            };

            GroupBox box = new GroupBox { Text = "本月水情摘要", Size = new Size(500, 200), Location = new Point(30, 90), Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblData = new Label { Text = "• 總取水量：1,250 m³\n• 總排放量：1,100 m³\n• 回收率：12%\n• 用藥狀態：正常", AutoSize = true, Location = new Point(20, 40) };
            box.Controls.Add(lblData);

            pnl.Controls.AddRange(new Control[] { lblTitle, box });
            main.Controls.Add(pnl);
            return main;
        }
    }
}
