using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_FireDashboard
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            Label t = new Label { Text = "🔥 消防安全管理儀表版", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 20) };
            p.Controls.Add(t);
            main.Controls.Add(p);
            return main;
        }
    }
}
