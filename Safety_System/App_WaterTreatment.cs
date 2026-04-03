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
        private const string TableName = "WaterTreatmentRecords";

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(0, 20, 0, 0) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            DataManager.InitTable(TableName, @"CREATE TABLE IF NOT EXISTS [WaterTreatmentRecords] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [廢水處理量] TEXT, [廢水進流量] TEXT, [納廢回收] TEXT, [貯存池水位] TEXT);");

            GroupBox box = new GroupBox { Text = "🌊 水處理操作紀錄", Dock = DockStyle.Fill, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            Button btnRead = new Button { Text = "讀取", Size = new Size(80, 35), BackColor = Color.LightBlue };
            btnRead.Click += (s, e) => RefreshGrid();

            Button btnSave = new Button { Text = "💾 儲存變更", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            btnSave.Click += (s, e) => {
                _dgv.EndEdit();
                DataTable dt = (DataTable)_dgv.DataSource;
                foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(TableName, r);
                MessageBox.Show("儲存成功");
            };

            flp.Controls.AddRange(new Control[] { new Label{Text="日期:"}, _dtpStart, new Label{Text="~"}, _dtpEnd, btnRead, btnSave });
            box.Controls.Add(flp);
            main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
        }
    }
}
