using System;
using System.Windows.Forms;
using System.Drawing;

namespace Safety_System
{
    public class App_WasteMonthly
    {
        public Control GetView()
        {
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Button btnExport = new Button { Text = "📤 導出月報表", Size = new Size(150, 40), BackColor = Color.LightGray };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White };
            dgv.Columns.Add("Month", "月份");
            dgv.Columns.Add("WasteType", "廢棄物代碼");
            dgv.Columns.Add("Weight", "產出重量(kg)");

            tlp.Controls.Add(btnExport, 0, 0);
            tlp.Controls.Add(dgv, 0, 1);
            return tlp;
        }
    }
}
