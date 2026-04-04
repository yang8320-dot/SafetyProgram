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

            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _dtpStart = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Short };
            Button bRead = new Button { Text = "讀取資料", Size = new Size(100, 35) };
            bRead.Click += (s, e) => RefreshGrid();

            Button bSave = new Button { Text = "💾 儲存變更", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => {
                _dgv.EndEdit();
                DataTable dt = (DataTable)_dgv.DataSource;
                foreach (DataRow r in dt.Rows) DataManager.UpsertRecord(DbName, TableName, r);
                MessageBox.Show("儲存完成！");
                RefreshGrid();
            };
            row1.Controls.AddRange(new Control[] { new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _dtpStart, _dtpEnd, bRead, bSave });

            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _txtNewColName = new TextBox { Width = 120 };
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(90, 35) };
            bAdd.Click += (s, e) => {
                if (string.IsNullOrEmpty(_txtNewColName.Text)) return;
                DataManager.AddColumn(DbName, TableName, _txtNewColName.Text);
                RefreshGrid();
            };

            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 120 };
            Button bRen = new Button { Text = "改名", Size = new Size(70, 35) };
            bRen.Click += (s, e) => {
                if (_cboColumns.SelectedItem == null || string.IsNullOrEmpty(_txtRenameCol.Text)) return;
                if (VerifyPassword()) {
                    DataManager.RenameColumn(DbName, TableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text);
                    RefreshGrid();
                }
            };

            Button bDel = new Button { Text = "刪除整列", Size = new Size(90, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDel.Click += (s, e) => {
                if (_dgv.CurrentRow == null || _dgv.CurrentRow.Cells["Id"].Value == DBNull.Value) return;
                if (VerifyPassword()) {
                    DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(_dgv.CurrentRow.Cells["Id"].Value));
                    RefreshGrid();
                }
            };
            row2.Controls.AddRange(new Control[] { new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(20, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDel });

            tlp.Controls.Add(row1, 0, 0);
            tlp.Controls.Add(row2, 0, 1);
            boxTop.Controls.Add(tlp);
            main.Controls.Add(boxTop, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            main.Controls.Add(_dgv, 0, 1);

            RefreshGrid();
            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private bool VerifyPassword() {
            // 🟢 修正：視窗高度從 240 增加到 270，防止按鈕被遮擋
            Form p = new Form { Width = 450, Height = 270, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "請輸入管理員密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
            
            // 🟢 修正：按鈕 Top 調整至 150，確保完整顯示
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            p.Controls.Add(lbl); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }
    }
}
