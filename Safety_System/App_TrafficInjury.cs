using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_TrafficInjury
    {
        public Control GetView()
        {
            GroupBox gb = new GroupBox { Text = "🚗 上下班交通意外紀錄 (交傷)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
            dgv.Columns.Add("Date", "日期");
            dgv.Columns.Add("Time", "時間");
            dgv.Columns.Add("Location", "事故路段");
            dgv.Columns.Add("Status", "處理進度");
            
            gb.Controls.Add(dgv);
            return gb;
        }
    }
}
