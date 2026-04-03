using System;
using System.Windows.Forms;
using System.Drawing;

namespace Safety_System
{
    public class App_HazardStats
    {
        public Control GetView()
        {
            GroupBox box = new GroupBox { Text = "☢️ 公共危險物品列管統計", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            dgv.Columns.Add("Date", "統計日期");
            dgv.Columns.Add("Type", "物品種類");
            dgv.Columns.Add("Qty", "存放數量");
            dgv.Columns.Add("Limit", "倍數限制");
            
            box.Controls.Add(dgv);
            return box;
        }
    }
}
