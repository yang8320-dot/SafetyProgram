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
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;

        private const string DbName = "Safety"; 
        private const string TableName = "NearMiss"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [NearMiss] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [地點] TEXT, [事件經過] TEXT, [提報人] TEXT, [改善措施] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "虛驚事件管理 (庫: Safety / 表: NearMiss)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 25, 10, 10) };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            // 第一行：查詢與儲存
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _dtpStart = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            _dtpEnd = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            row1.Controls.AddRange(new Control[] { new Label { Text = "日期:", AutoSize = true }, _dtpStart, _dtpEnd });
            
            Button btnRead = new Button { Text = "讀取", Size = new Size(100, 35), BackColor = Color.LightBlue };
            btnRead.Click += (s, e) => RefreshGrid();
            row1.Controls.Add(btnRead);

            Button btnSave = new Button { Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            btnSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(DbName, TableName, r); MessageBox.Show("儲存成功"); RefreshGrid(); };
            row1.Controls.Add(btnSave);

            // 第二行：欄位管理 (活化變數，消除警告)
            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 10, 0, 5) };
            _txtNewColName = new TextBox { Width = 120 };
            Button btnAdd = new Button { Text = "新增欄", Size = new Size(80, 35) };
            btnAdd.Click += (s, e) => { if(!string.IsNullOrEmpty(_txtNewColName.Text)) { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text); RefreshGrid(); } };
            
            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 120 };
            Button btnRen = new Button { Text = "✏️ 改名", Size = new Size(80, 35), BackColor = Color.LightYellow };
            btnRen.Click += (s, e) => { if(VerifyPassword()) { DataManager.RenameColumn(DbName, TableName, _cboColumns.Text, _txtRenameCol.Text); RefreshGrid(); } };
            
            Button btnDel = new Button { Text = "🗑️ 刪除列", Size = new Size(100, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            btnDel.Click += (s, e) => { if (_dgv.CurrentRow != null && MessageBox.Show("確定刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) { var id = _dgv.CurrentRow.Cells["Id"].Value; if (id != DBNull.Value) DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(id)); RefreshGrid(); } };
            
            row2.Controls.AddRange(new Control[] { new Label { Text = "欄位:" }, _txtNewColName, btnAdd, _cboColumns, _txtRenameCol, btnRen, btnDel });

            tlp.Controls.Add(row1, 0, 0); tlp.Controls.Add(row2, 0, 1);
            boxTop.Controls.Add(tlp); main.Controls.Add(boxTop, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 450, Height = 240, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 90, Font = new Font("UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 145, Width = 120, Height = 40 };
            p.Controls.AddRange(new Control[] { new Label { Text = "請輸入管理員密碼:", Left = 30, Top = 30, Font = new Font("UI", 14F) }, t, b });
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }
    }
}
