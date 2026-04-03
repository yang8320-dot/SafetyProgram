using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_AirDashboard
    {
        public Control GetView()
        {
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            Label lbl = new Label { Text = "☁️ 空氣汙染防制監測儀表版", Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 30) };
            
            ProgressBar pb = new ProgressBar { Location = new Point(35, 100), Width = 300, Height = 30, Value = 75 };
            Label lblStatus = new Label { Text = "本季申報進度: 75%", Location = new Point(35, 140), AutoSize = true };
            
            p.Controls.AddRange(new Control[] { lbl, pb, lblStatus });
            return p;
        }
    }
}
