using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_DischargeData
    {
        private DataGridView _dgv;
        private const string TableName = "WaterDischargeRecords";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [WaterDischargeRecords] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [COD值] TEXT, [SS值] TEXT, [pH值] TEXT, [排放量] TEXT);");

            GroupBox gb = new GroupBox { Text = "📉 納管排放水質數據", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button b = new Button { Text = "💾 儲存數據", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White, Location = new Point(20, 30) };
            b.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("數據已更新"); };
            
            gb.Controls.Add(b);
            main.Controls.Add(gb, 0, 0);
            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
