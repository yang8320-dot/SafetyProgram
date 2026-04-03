using System;
using System.Windows.Forms;
using System.Drawing;

namespace Safety_System
{
    public class App_FireResponsible
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(10) };
            Button btnAdd = new Button { Text = "新增責任人", Size = new Size(120, 35), BackColor = Color.LightBlue };
            Button btnSave = new Button { Text = "儲存變更", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            
            flp.Controls.AddRange(new Control[] { btnAdd, btnSave });
            
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
            dgv.Columns.Add("Area", "管轄區域");
            dgv.Columns.Add("Name", "負責人姓名");
            dgv.Columns.Add("Ext", "分機");

            main.Controls.Add(flp, 0, 0);
            main.Controls.Add(dgv, 0, 1);
            return main;
        }
    }
}
