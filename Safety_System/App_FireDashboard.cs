using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_FireDashboard
    {
        public Control GetView()
        {
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, Padding = new Padding(20) };
            Label lblTitle = new Label { 
                Text = "🔥 消防安全監控儀表版", 
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold), 
                AutoSize = true, Location = new Point(20, 20) 
            };
            
            // 範例統計方塊
            GroupBox boxStats = new GroupBox { Text = "本月概況", Size = new Size(400, 150), Location = new Point(20, 80), Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblInfo = new Label { Text = "• 設備檢查完成度: 95%\n• 火源責任人簽署: 12/12\n• 公危物品申報狀態: 已完成", AutoSize = true, Location = new Point(20, 40) };
            
            boxStats.Controls.Add(lblInfo);
            pnl.Controls.AddRange(new Control[] { lblTitle, boxStats });
            return pnl;
        }
    }
}
