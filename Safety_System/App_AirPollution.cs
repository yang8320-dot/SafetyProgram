using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_AirPollution
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "空污排放申報管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10) };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Button btnAdd = new Button { Text = "新增申報紀錄", Size = new Size(130, 35), BackColor = Color.LightBlue };
            Button btnSave = new Button { Text = "💾 儲存數據", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            
            flp.Controls.AddRange(new Control[] { btnAdd, btnSave });
            boxTop.Controls.Add(flp);

            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill 
            };
            dgv.Columns.Add("Date", "申報日期");
            dgv.Columns.Add("Type", "排放種類");
            dgv.Columns.Add("Value", "排放量(m³)");
            dgv.Columns.Add("Status", "申報狀態");

            main.Controls.Add(boxTop, 0, 0);
            main.Controls.Add(dgv, 0, 1);

            return main;
        }
    }
}
