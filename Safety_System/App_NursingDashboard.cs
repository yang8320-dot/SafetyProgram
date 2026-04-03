using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_NursingDashboard
    {
        public Control GetView()
        {
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.SeaShell };
            Label lbl = new Label { Text = "🏥 職場健康與護理儀表版", Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 30) };
            p.Controls.Add(lbl);
            return p;
        }
    }
}
