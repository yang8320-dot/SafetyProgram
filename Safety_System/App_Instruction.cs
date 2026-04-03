using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_Instruction
    {
        public Control GetView()
        {
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };

            GroupBox gb = new GroupBox {
                Text = "系統管理員說明與配置",
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold),
                Location = new Point(50, 50),
                Size = new Size(600, 200)
            };

            Label lblPwd = new Label {
                Text = "⚠️ 高權限操作授權\n\n水資源表單管理 (修改與刪除欄位等進階操作)\n通用授權密碼：tces",
                Font = new Font("Microsoft JhengHei UI", 14F),
                ForeColor = Color.DarkRed,
                Location = new Point(30, 60),
                AutoSize = true
            };

            gb.Controls.Add(lblPwd);
            pnl.Controls.Add(gb);

            return pnl;
        }
    }
}
