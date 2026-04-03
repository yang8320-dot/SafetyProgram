using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_AirDashboard
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            Label lbl = new Label { Text = "☁️ 空氣汙染防治監測儀表版", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 20) };
            pnl.Controls.Add(lbl);
            main.Controls.Add(pnl);
            return main;
        }
    }
}
