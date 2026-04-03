using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WasteDashboard
    {
        public Control GetView()
        {
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, Padding = new Padding(30) };
            Label t = new Label { Text = "♻️ 廢棄物管理儀表版", Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), AutoSize = true };
            p.Controls.Add(t);
            return p;
        }
    }
}
