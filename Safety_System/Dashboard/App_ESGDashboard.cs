/// FILE: Safety_System/App_ESGDashboard.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ESGDashboard
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
                Text = "🌱 ESG 永續發展管理看板", 
                Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), 
                ForeColor = Color.DarkOliveGreen,
                AutoSize = true, 
                Location = new Point(30, 20) 
            };

            Label lblSub = new Label 
            { 
                Text = "此區域為 ESG (環境、社會、公司治理) 專屬數據看板。\n未來可在此擴充碳排放量圖表、節能減碳專案進度追蹤與永續報告書數據可視化。", 
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
