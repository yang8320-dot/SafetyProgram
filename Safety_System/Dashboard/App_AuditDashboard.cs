using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_AuditDashboard
    {
        public Control GetView()
        {
            // 🟢 增加頂部 Padding 防止被遮擋
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            
            Label lblTitle = new Label { 
                Text = "🛡️ 稽核資料查詢", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                AutoSize = true, Location = new Point(30, 20) 
            };

            GroupBox box = new GroupBox { Text = "系統提示", Size = new Size(500, 200), Location = new Point(30, 80), Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblDesc = new Label { Text = "本看板為稽核資料查詢，看板建置中...", ForeColor = Color.DimGray, AutoSize = true, Location = new Point(20, 50) };
            box.Controls.Add(lblDesc);

            pnl.Controls.AddRange(new Control[] { lblTitle, box });
            main.Controls.Add(pnl);
            return main;
        }
    }
}
