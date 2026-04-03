using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterDashboard
    {
        public Control GetView()
        {
            Panel pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke };
            
            Label lbl = new Label {
                Text = "本頁面對應 App_WaterDashboard.cs, 資料尚在建立中\n\n(未來功能：水資源數據視覺化圖表、總覽與統計)",
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Italic),
                ForeColor = Color.DarkGray,
                AutoSize = true,
                Location = new Point(50, 50)
            };
            
            pnl.Controls.Add(lbl);
            
            return pnl;
        }
    }
}
