using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemQuickSearch
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            Label t = new Label { Text = "🔍 化學品快查系統", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), AutoSize = true, Location = new Point(30, 20) };
            Label sub = new Label { Text = "請輸入 CAS No. 或化學品名稱進行跨表快速檢索...", Font = new Font("Microsoft JhengHei UI", 14F), AutoSize = true, Location = new Point(35, 70), ForeColor = Color.Gray };
            
            p.Controls.Add(t); 
            p.Controls.Add(sub); 
            main.Controls.Add(p); 
            return main;
        }
    }
}
