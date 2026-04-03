using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_NearMiss
    {
        private DataGridView _dgv;
        private const string TableName = "NearMissRecords";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [NearMissRecords] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [事件描述] TEXT, [潛在危險] TEXT, [預防措施] TEXT);");

            GroupBox gb = new GroupBox { Text = "⚡ 虛驚事件提報", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button bSave = new Button { Text = "💾 儲存所有紀錄", Size = new Size(150, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Location = new Point(20, 30) };
            bSave.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("已更新"); };
            
            gb.Controls.Add(bSave);
            main.Controls.Add(gb, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
