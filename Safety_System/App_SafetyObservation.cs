using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyObservation
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { 
                Text = "🔍 安全觀查 (Safety Observation)", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                AutoSize = true, 
                Padding = new Padding(10) 
            };
            
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            Button btnAdd = new Button { Text = "新增觀查項", Size = new Size(120, 35), BackColor = Color.LightBlue };
            Button btnSave = new Button { Text = "💾 儲存數據", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            flp.Controls.AddRange(new Control[] { btnAdd, btnSave });
            boxTop.Controls.Add(flp);

            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill 
            };
            dgv.Columns.Add("Date", "觀查日期");
            dgv.Columns.Add("Location", "區域/地點");
            dgv.Columns.Add("Category", "觀查類別"); // 例如：不安全行為、不安全環境
            dgv.Columns.Add("Observer", "觀查人");
            dgv.Columns.Add("Description", "描述與建議");

            main.Controls.Add(boxTop, 0, 0);
            main.Controls.Add(dgv, 0, 1);

            return main;
        }
    }
}
