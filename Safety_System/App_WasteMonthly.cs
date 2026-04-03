using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WasteMonthly
    {
        private DataGridView _dgv;
        private const string TableName = "WasteMonthlyRecords";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [WasteMonthlyRecords] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [月份] TEXT, [廢棄物代碼] TEXT, [產出重量(kg)] TEXT, [清運廠商] TEXT);");

            GroupBox gb = new GroupBox { Text = "📦 廢棄物產出月報表", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button b = new Button { Text = "💾 儲存月報", Size = new Size(120, 35), BackColor = Color.Sienna, ForeColor = Color.White, Location = new Point(20, 30) };
            b.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("月報表已儲存"); };
            
            gb.Controls.Add(b);
            main.Controls.Add(gb, 0, 0);
            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
