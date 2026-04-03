using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_AirPollution
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewCol; 

        private const string DbName = "Air"; 
        private const string TableName = "AirPollution"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [AirPollution] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [排放口] TEXT, [污染物] TEXT, [濃度] TEXT, [狀態] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { Text = "空污申報紀錄 (庫: Air / 表: AirPollution)", Dock = DockStyle.Fill, Font = new Font("UI", 12F), AutoSize = true };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(5) };
            
            _dtpStart = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Short };
            Button bRead = new Button { Text = "讀取", Size = new Size(80, 35), BackColor = Color.LightBlue };
            bRead.Click += (s, e) => RefreshGrid();
            
            _txtNewCol = new TextBox { Width = 120, Margin = new Padding(20, 5, 3, 3) };
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(90, 35) };
            bAdd.Click += (s, e) => { if(!string.IsNullOrEmpty(_txtNewCol.Text)) { DataManager.AddColumn(DbName, TableName, _txtNewCol.Text); RefreshGrid(); } };
            
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(80, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) DataManager.UpsertRecord(DbName, TableName, r); MessageBox.Show("儲存完成"); RefreshGrid(); };

            flp.Controls.AddRange(new Control[] { new Label { Text = "區間:" }, _dtpStart, _dtpEnd, bRead, _txtNewCol, bAdd, bSave });
            box.Controls.Add(flp); main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
        }
    }
}
