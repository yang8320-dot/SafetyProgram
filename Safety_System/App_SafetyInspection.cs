using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyInspection
    {
        private DataGridView _dgv;
        private const string TableName = "SafetyInspection";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 初始化資料表
            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [SafetyInspection] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [巡檢區域] TEXT, [檢查結果] TEXT, [改善措施] TEXT, [負責人] TEXT);");

            GroupBox box = new GroupBox { Text = "巡檢記錄管理", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            Button bRead = new Button { Text = "讀取", Size = new Size(100, 35), BackColor = Color.LightBlue };
            bRead.Click += (s, e) => RefreshGrid();
            
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => { _dgv.EndEdit(); foreach (DataRow r in ((DataTable)_dgv.DataSource).Rows) DataManager.UpsertRecord(TableName, r); MessageBox.Show("儲存成功"); };

            flp.Controls.AddRange(new Control[] { bRead, bSave });
            box.Controls.Add(flp);
            main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);

            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(TableName, "日期", DateTime.Now.AddDays(-30).ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd"));
        }
    }
}
