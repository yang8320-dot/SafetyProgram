using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_SafetyInspection
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;

        private const string DbName = "Safety"; 
        private const string TableName = "SafetyInspection"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [SafetyInspection] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [巡檢區域] TEXT, 
                [檢查項目] TEXT, 
                [檢查結果] TEXT, 
                [缺失描述] TEXT, 
                [改善措施] TEXT, 
                [負責人] TEXT, 
                [狀態] TEXT);");

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { 
                Text = "巡檢記錄管理 (庫: " + DbName + " / 表: " + TableName + ")", 
                Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 25, 10, 10)
            };

            TableLayoutPanel tlpControls = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            // 第 1 行
            FlowLayoutPanel flpRow1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            flpRow1.Controls.Add(new Label { Text = "日期起:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _dtpStart = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Margin = new Padding(3, 5, 3, 3) };
            _dtpStart.Value = DateTime.Now.AddDays(-30); 
            flpRow1.Controls.Add(_dtpStart);
            flpRow1.Controls.Add(new Label { Text = "日期迄:", AutoSize = true, Margin = new Padding(15, 10, 3, 0) });
            _dtpEnd = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Margin = new Padding(3, 5, 3, 3) };
            flpRow1.Controls.Add(_dtpEnd);
            Button btnRead = new Button { Text = "讀取資料", Size = new Size(110, 35), BackColor = Color.LightBlue };
            btnRead.Click += (s, e) => RefreshGrid();
            flpRow1.Controls.Add(btnRead);
            Button btnSave = new Button { Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            btnSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(DbName, TableName, r); MessageBox.Show("儲存成功"); RefreshGrid(); };
            flpRow1.Controls.Add(btnSave);

            // 第 2 行
            FlowLayoutPanel flpRow2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 10, 0, 5) };
            flpRow2.Controls.Add(new Label { Text = "新增欄位:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _txtNewColName = new TextBox { Width = 150 }; flpRow2.Controls.Add(_txtNewColName);
            Button btnAddCol = new Button { Text = "確認新增", Size = new Size(100, 35), BackColor = Color.LightGray };
            btnAddCol.Click += (s, e) => { if (!string.IsNullOrWhiteSpace(_txtNewColName.Text)) { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text.Trim()); _txtNewColName.Clear(); RefreshGrid(); } };
            flpRow2.Controls.Add(btnAddCol);
            flpRow2.Controls.Add(new Label { Text = "管理:", AutoSize = true, Margin = new Padding(25, 10, 3, 0) });
            _cboColumns = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList }; flpRow2.Controls.Add(_cboColumns);
            _txtRenameCol = new TextBox { Width = 140 }; flpRow2.Controls.Add(_txtRenameCol);
            Button btnRen = new Button { Text = "✏️", Size = new Size(50, 35), BackColor = Color.LightYellow };
            btnRen.Click += (s, e) => { if(VerifyPassword()) { DataManager.RenameColumn(DbName, TableName, _cboColumns.Text, _txtRenameCol.Text); RefreshGrid(); } };
            flpRow2.Controls.Add(btnRen);
            Button btnDelR = new Button { Text = "🗑️ 刪除選取", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            btnDelR.Click += (s, e) => { if (_dgv.CurrentRow != null && MessageBox.Show("確定刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) { var id = _dgv.CurrentRow.Cells["Id"].Value; if (id != DBNull.Value) DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(id)); RefreshGrid(); } };
            flpRow2.Controls.Add(btnDelR);

            tlpControls.Controls.Add(flpRow1, 0, 0); tlpControls.Controls.Add(flpRow2, 0, 1);
            boxTop.Controls.Add(tlpControls); mainLayout.Controls.Add(boxTop, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, AllowUserToAddRows = true };
            _dgv.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.V) PasteClipboard(); };
            mainLayout.Controls.Add(_dgv, 0, 1);

            return mainLayout;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn col in _dgv.Columns) if (col.Name != "Id" && col.Name != "日期") _cboColumns.Items.Add(col.Name);
        }

        private bool VerifyPassword() {
            Form prompt = new Form() { Width = 450, Height = 240, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "請輸入管理員密碼：", AutoSize = true, Font = new Font("UI", 14F) };
            TextBox txt = new TextBox() { Left = 30, Top = 95, Width = 370, PasswordChar = '*', Font = new Font("UI", 14F) };
            Button btn = new Button() { Text = "確認", Left = 280, Top = 145, Width = 120, Height = 40, DialogResult = DialogResult.OK };
            prompt.Controls.Add(lbl); prompt.Controls.Add(txt); prompt.Controls.Add(btn);
            return prompt.ShowDialog() == DialogResult.OK && txt.Text == "tces";
        }

        private void PasteClipboard() {
            try {
                string text = Clipboard.GetText(); if (string.IsNullOrEmpty(text)) return;
                string[] lines = text.Split('\n'); int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex;
                DataTable dt = (DataTable)_dgv.DataSource;
                foreach (string line in lines) {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                    string[] cells = line.Split('\t');
                    for (int i = 0; i < cells.Length; i++) if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly) _dgv[c + i, r].Value = cells[i].Trim();
                    r++;
                }
            } catch { }
        }
    }
}
