using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyInspection
    {
        public Control GetView()
        {
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "📋 巡檢記錄管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10) };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Button btnAdd = new Button { Text = "新增巡檢項", Size = new Size(120, 35), BackColor = Color.LightBlue };
            Button btnSave = new Button { Text = "儲存記錄", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            
            flp.Controls.AddRange(new Control[] { btnAdd, btnSave });
            boxTop.Controls.Add(flp);

            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill 
            };
            dgv.Columns.Add("Date", "日期");
            dgv.Columns.Add("Area", "區域");
            dgv.Columns.Add("Result", "檢查結果");
            dgv.Columns.Add("Inspector", "巡檢人");

            tlp.Controls.Add(boxTop, 0, 0);
            tlp.Controls.Add(dgv, 0, 1);
            return tlp;
        }
    }
}
