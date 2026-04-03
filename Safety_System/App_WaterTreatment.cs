using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterTreatment
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewCol, _txtRenameCol;
        private ComboBox _cboCols;
        private const string TableName = "WaterMeterReadings"; // 🟢 指定此選單的資料表

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 初始化資料表結構
            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [WaterMeterReadings] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [廢水處理量] TEXT, [廢水進流量] TEXT, 
                [納廢回收6吋] TEXT, [雙介質A] TEXT, [雙介質B] TEXT, [貯存池] TEXT, [軟水A] TEXT, [軟水B] TEXT, [軟水C] TEXT);");

            // --- UI 佈局 (採用您要求的兩行動態排版) ---
            GroupBox box = new GroupBox { Text = "水處理記錄管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            row1.Controls.Add(new Label { Text = "日期起:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _dtpStart = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            _dtpStart.Value = DateTime.Now.AddDays(-30);
            row1.Controls.Add(_dtpStart);
            row1.Controls.Add(new Label { Text = "日期迄:", AutoSize = true, Margin = new Padding(15, 10, 3, 0) });
            _dtpEnd = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            row1.Controls.Add(_dtpEnd);
            
            Button bRead = new Button { Text = "讀取", Size = new Size(100, 35), BackColor = Color.LightBlue };
            bRead.Click += (s, e) => RefreshGrid();
            row1.Controls.Add(bRead);
            
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => {
                _dgv.EndEdit();
                DataTable dt = (DataTable)_dgv.DataSource;
                foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(TableName, r);
                MessageBox.Show("儲存成功"); RefreshGrid();
            };
            row1.Controls.Add(bSave);

            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            row2.Controls.Add(new Label { Text = "新增欄位:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _txtNewCol = new TextBox { Width = 120 }; row2.Controls.Add(_txtNewCol);
            Button bAdd = new Button { Text = "確認", Size = new Size(80, 35) };
            bAdd.Click += (s, e) => { DataManager.AddColumn(TableName, _txtNewCol.Text); RefreshGrid(); };
            row2.Controls.Add(bAdd);
            
            Button bDelR = new Button { Text = "🗑️ 刪除選取資料", Size = new Size(150, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelR.Click += (s, e) => {
                if (_dgv.CurrentRow != null && MessageBox.Show("確定刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    var id = _dgv.CurrentRow.Cells["Id"].Value;
                    if (id != DBNull.Value) DataManager.DeleteRecord(TableName, Convert.ToInt32(id));
                    RefreshGrid();
                }
            };
            row2.Controls.Add(bDelR);

            tlp.Controls.Add(row1, 0, 0); tlp.Controls.Add(row2, 0, 1);
            box.Controls.Add(tlp); main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            main.Controls.Add(_dgv, 0, 1);

            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
        }
    }
}
