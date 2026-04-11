using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemDashboard
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            Label t = new Label { Text = "🧪 化學品管理看板", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 20) };
            Label sub = new Label { Text = "此為獨立開發區域，未來可在此放置化學品用量圖表與高風險物質警示。", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.DimGray, AutoSize = true, Location = new Point(35, 75) };
            
            p.Controls.Add(t);
            p.Controls.Add(sub);
            main.Controls.Add(p);
            return main;
        }
    }
}
