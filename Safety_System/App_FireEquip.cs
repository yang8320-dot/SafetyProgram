using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_FireEquip
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "消防器材清冊與巡檢紀錄", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10) };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Button btnAdd = new Button { Text = "新增器材", Size = new Size(100, 35), BackColor = Color.LightBlue };
            Button btnExport = new Button { Text = "📊 導出清冊", Size = new Size(100, 35), BackColor = Color.LightGray };
            
            flp.Controls.AddRange(new Control[] { btnAdd, btnExport });
            boxTop.Controls.Add(flp);

            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells 
            };
            dgv.Columns.Add("Id", "設備編號");
            dgv.Columns.Add("Location", "放置位置");
            dgv.Columns.Add("Type", "器材種類");
            dgv.Columns.Add("ExpDate", "有效/換藥日期");
            dgv.Columns.Add("Result", "上次檢查結果");

            main.Controls.Add(boxTop, 0, 0);
            main.Controls.Add(dgv, 0, 1);

            return main;
        }
    }
}
