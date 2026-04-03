using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_YearlyReport
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            Label t = new Label { Text = "🏆 系統年度績效報表", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 20) };
            p.Controls.Add(t);
            main.Controls.Add(p);
            return main;
        }
    }
}
