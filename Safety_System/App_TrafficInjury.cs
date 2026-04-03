using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_TrafficInjury
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName;

        private const string DbName = "Safety"; 
        private const string TableName = "TrafficInjury"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [TrafficInjury] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [時間] TEXT, 
                [姓名] TEXT, 
                [事故地點] TEXT, 
                [事故經過] TEXT, 
                [受傷狀況] TEXT, 
                [假別] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { Text = "交通意外紀錄 (庫: " + DbName + ")", Dock = DockStyle.Fill, Font = new Font("UI", 12F), AutoSize = true };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            Button bRead = new Button { Text = "讀取", Size = new Size(80, 35) };
            bRead.Click += (s, e) => _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(80, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) DataManager.UpsertRecord(DbName, TableName, r); MessageBox.Show("儲存完成"); };

            flp.Controls.AddRange(new Control[] { new Label { Text = "日期範圍:" }, _dtpStart, _dtpEnd, bRead, bSave });
            box.Controls.Add(flp); main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }
    }
}
