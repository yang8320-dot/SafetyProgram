using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_NearMiss
    {
        public Control GetView()
        {
            GroupBox gb = new GroupBox { Text = "⚡ 虛驚事件 (Near Miss) 提報單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
            dgv.Columns.Add("Date", "發生日期");
            dgv.Columns.Add("Location", "地點");
            dgv.Columns.Add("Desc", "事件描述");
            dgv.Columns.Add("Action", "改善措施");
            
            gb.Controls.Add(dgv);
            return gb;
        }
    }
}
