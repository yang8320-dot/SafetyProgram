using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyObservation
    {
        private DataGridView _dgv;
        private const string TableName = "SafetyObservation";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [SafetyObservation] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [觀查人員] TEXT, [行為描述] TEXT, [安全類別] TEXT, [建議] TEXT);");

            GroupBox gb = new GroupBox { Text = "🔍 安全行為觀查 (BBS)", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button btn = new Button { Text = "💾 儲存資料", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Location = new Point(20, 30) };
            btn.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("儲存成功"); };
            
            gb.Controls.Add(btn);
            main.Controls.Add(gb, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
