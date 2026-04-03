using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_HazardStats
    {
        private DataGridView _dgv;
        private const string TableName = "HazardousMaterialStats";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [HazardousMaterialStats] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [物品名稱] TEXT, [存量] TEXT, [管制量倍數] TEXT, [存放地點] TEXT);");

            GroupBox gb = new GroupBox { Text = "☢️ 公共危險物品列管統計", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button b = new Button { Text = "💾 儲存統計", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White, Location = new Point(20, 30) };
            b.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("統計數據已更新"); };
            
            gb.Controls.Add(b);
            main.Controls.Add(gb, 0, 0);
            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
