using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WasteMonthly
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewCol, _txtRenameCol;
        private ComboBox _cboCols;

        private const string DbName = "Waste"; 
        private const string TableName = "WasteMonthly"; 

        public Control GetView()
        {
            // 1. 初始化資料表
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WasteMonthly] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [廢棄物代碼] TEXT, 
                [廢棄物名稱] TEXT, 
                [產出量_kg] TEXT, 
                [清理商] TEXT, 
                [聯單編號] TEXT, 
                [備註] TEXT);");

            // 2. 主排版
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { Text = "廢棄物生產月報管理 (庫: Waste)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            // 第一行：查詢與儲存
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            row1.Controls.Add(new Label { Text = "日期:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _dtpStart = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            _dtpEnd = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            row1.Controls.Add(_dtpStart); row1.Controls.Add(_dtpEnd);
            
            Button bRead = new Button { Text = "讀取", Size = new Size(100, 35), BackColor = Color.LightBlue };
            bRead.Click += (s, e) => RefreshGrid();
            row1.Controls.Add(bRead);

            Button bSave = new Button { Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => {
                _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource;
                if (dt == null) return;
                foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(DbName, TableName, r);
                MessageBox.Show("儲存成功"); RefreshGrid();
            };
            row1.Controls.Add(bSave);

            // 第二行：管理
            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            row2.Controls.Add(new Label { Text = "新增:", AutoSize = true });
            _txtNewCol = new TextBox { Width = 120 }; row2.Controls.Add(_txtNewCol);
            Button bAdd = new Button { Text = "確認", Size = new Size(80, 35) };
            bAdd.Click += (s, e) => { DataManager.AddColumn(DbName, TableName, _txtNewCol.Text); RefreshGrid(); };
            row2.Controls.Add(bAdd);
            
            Button bDelR = new Button { Text = "🗑️ 刪除列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelR.Click += (s, e) => {
                if (_dgv.CurrentRow != null && MessageBox.Show("確定刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    var id = _dgv.CurrentRow.Cells["Id"].Value;
                    if (id != DBNull.Value) DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(id));
                    RefreshGrid();
                }
            };
            row2.Controls.Add(bDelR);

            tlp.Controls.Add(row1, 0, 0); tlp.Controls.Add(row2, 0, 1);
            box.Controls.Add(tlp); main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true };
            _dgv.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.V) PasteClipboard(); };
            main.Controls.Add(_dgv, 0, 1);

            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
        }

        private void PasteClipboard() {
            try {
                string t = Clipboard.GetText(); string[] lines = t.Split('\n');
                DataTable dt = (DataTable)_dgv.DataSource;
                int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex;
                foreach (var l in lines) {
                    if (string.IsNullOrWhiteSpace(l)) continue;
                    if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                    string[] cells = l.Split('\t');
                    for (int i = 0; i < cells.Length; i++) if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly) _dgv[c + i, r].Value = cells[i].Trim();
                    r++;
                }
            } catch { }
        }
    }
}
