using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_NearMiss
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol; // 🟢 已使用
        private ComboBox _cboColumns; // 🟢 已使用

        private const string DbName = "Safety"; 
        private const string TableName = "NearMiss"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [NearMiss] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [發生地點] TEXT, [事件描述] TEXT, [提報人] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { Text = "虛驚事件管理", Dock = DockStyle.Fill, Font = new Font("UI", 12F), AutoSize = true };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            Button bRead = new Button { Text = "讀取", Size = new Size(80, 35) };
            bRead.Click += (s, e) => RefreshGrid();
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(80, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) DataManager.UpsertRecord(DbName, TableName, r); RefreshGrid(); };
            row1.Controls.AddRange(new Control[] { _dtpStart, _dtpEnd, bRead, bSave });

            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _txtNewColName = new TextBox { Width = 100 };
            Button bAdd = new Button { Text = "新增欄", Size = new Size(70, 35) };
            bAdd.Click += (s, e) => { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text); RefreshGrid(); };
            
            _cboColumns = new ComboBox { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 100 };
            Button bRen = new Button { Text = "✏️", Size = new Size(40, 35) };
            bRen.Click += (s, e) => { if(VerifyPassword()) DataManager.RenameColumn(DbName, TableName, _cboColumns.Text, _txtRenameCol.Text); RefreshGrid(); };
            
            row2.Controls.AddRange(new Control[] { _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen });

            tlp.Controls.Add(row1, 0, 0); tlp.Controls.Add(row2, 0, 1);
            box.Controls.Add(tlp); main.Controls.Add(box, 0, 0);
            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if(c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 300, Height = 150, Text = "驗證" };
            TextBox t = new TextBox { PasswordChar = '*', Width = 200, Left = 40, Top = 20 };
            Button b = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 100, Top = 60 };
            p.Controls.AddRange(new Control[] { t, b });
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }
    }
}
