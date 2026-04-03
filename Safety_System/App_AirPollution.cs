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
            // 初始化資料表
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [AirPollution] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [排放口編號] TEXT, 
                [污染物種類] TEXT, 
                [排放濃度] TEXT, 
                [排放量] TEXT, 
                [申報狀態] TEXT,
                [備註] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { Text = "空污排放申報紀錄 (庫: Air)", Dock = DockStyle.Fill, Font = new Font("UI", 12F), AutoSize = true };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            _dtpStart = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            _dtpEnd = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            Button bRead = new Button { Text = "讀取", Size = new Size(90, 35), BackColor = Color.LightBlue };
            bRead.Click += (s, e) => RefreshGrid();
            
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(90, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => {
                _dgv.EndEdit();
                DataTable dt = (DataTable)_dgv.DataSource;
                foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(DbName, TableName, r);
                MessageBox.Show("儲存成功"); RefreshGrid();
            };

            flp.Controls.AddRange(new Control[] { new Label { Text = "日期:" }, _dtpStart, _dtpEnd, bRead, bSave });
            box.Controls.Add(flp); main.Controls.Add(box, 0, 0);

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
