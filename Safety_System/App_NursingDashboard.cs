using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_NursingDashboard
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.MistyRose };
            Label t = new Label { Text = "🏥 職場健康管理儀表版", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 20) };
            p.Controls.Add(t);
            main.Controls.Add(p);
            return main;
        }
    }
}
