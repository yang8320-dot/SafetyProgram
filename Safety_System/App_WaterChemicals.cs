using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterChemicals
    {
        private DataGridView _dgv;
        private const string TableName = "WaterChemicalRecords";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [WaterChemicalRecords] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [藥劑名稱] TEXT, [投藥量(kg)] TEXT, [操作人員] TEXT);");

            GroupBox box = new GroupBox { Text = "🧪 水處理用藥紀錄", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button btnSave = new Button { Text = "💾 儲存存檔", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Location = new Point(20, 30) };
            btnSave.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("已儲存"); };
            
            box.Controls.Add(btnSave);
            main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
