using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WorkInjury
    {
        private DataGridView _dgv;
        private const string TableName = "WorkInjuryRecords";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [WorkInjuryRecords] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [事故日期] TEXT, [受傷人員] TEXT, [受傷部位] TEXT, [損失日數] TEXT, [原因分析] TEXT);");

            GroupBox gb = new GroupBox { Text = "🩹 工傷事件紀錄表", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button btn = new Button { Text = "💾 儲存存檔", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White, Location = new Point(20, 30) };
            btn.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("工傷資料已儲存"); };
            
            gb.Controls.Add(btn);
            main.Controls.Add(gb, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
