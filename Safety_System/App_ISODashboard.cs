/// FILE: Safety_System/App_ISODashboard.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ISODashboard
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel 
            { 
                Dock = DockStyle.Fill, 
                Padding = new Padding(0, 20, 0, 0) 
            };
            
            Panel pnl = new Panel 
            { 
                Dock = DockStyle.Fill, 
                BackColor = Color.WhiteSmoke 
            };
            
            Label lblTitle = new Label 
            { 
                Text = "📐 ISO 14001 環境管理系統看板", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                ForeColor = Color.MidnightBlue,
                AutoSize = true, 
                Location = new Point(30, 20) 
            };

            Label lblSub = new Label 
            { 
                Text = "此區域為 ISO 14001 專屬管理區域。\n可在此追蹤環境管理目標、稽核計畫與持續改善(PDCA)之績效指標。", 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                ForeColor = Color.DimGray, 
                AutoSize = true, 
                Location = new Point(35, 75) 
            };

            pnl.Controls.Add(lblTitle);
            pnl.Controls.Add(lblSub);
            main.Controls.Add(pnl);
            
            return main;
        }
    }
}
