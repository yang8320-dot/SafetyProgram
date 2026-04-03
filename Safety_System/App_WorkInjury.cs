using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WorkInjury
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;

        private const string DbName = "Safety"; 
        private const string TableName = "WorkInjury"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WorkInjury] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [單位] TEXT, 
                [姓名] TEXT, 
                [受傷部位] TEXT, 
                [事故原因] TEXT, 
                [損失日數] TEXT, 
                [改善對策] TEXT);");

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "工傷事件紀錄 (庫: " + DbName + " / 表: " + TableName + ")", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 25, 10, 10) };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            row1.Controls.Add(new Label { Text = "日期:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });
            _dtpStart = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            _dtpEnd = new DateTimePicker { Width = 180, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            row1.Controls.Add(_dtpStart); row1.Controls.Add(_dtpEnd);
            Button bRead = new Button { Text = "讀取", Size = new Size(100, 35), BackColor = Color.LightBlue }; 
            bRead.Click += (s, e) => RefreshGrid();
            row1.Controls.Add(bRead);
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) if (r.RowState != DataRowState.Deleted) DataManager.UpsertRecord(DbName, TableName, r); MessageBox.Show("儲存成功"); RefreshGrid(); };
            row1.Controls.Add(bSave);

            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            row2.Controls.Add(new Label { Text = "新增欄位:", AutoSize = true });
            _txtNewColName = new TextBox { Width = 120 }; row2.Controls.Add(_txtNewColName);
            Button bAdd = new Button { Text = "確認", Size = new Size(80, 35) };
            bAdd.Click += (s, e) => { DataManager.AddColumn(DbName, TableName, _txtNewColName.Text); RefreshGrid(); };
            row2.Controls.Add(bAdd);
            Button bDelR = new Button { Text = "🗑️ 刪除選取", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelR.Click += (s, e) => { if (_dgv.CurrentRow != null && MessageBox.Show("確定刪除?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) { var id = _dgv.CurrentRow.Cells["Id"].Value; if (id != DBNull.Value) DataManager.DeleteRecord(DbName, TableName, Convert.ToInt32(id)); RefreshGrid(); } };
            row2.Controls.Add(bDelR);

            tlp.Controls.Add(row1, 0, 0); tlp.Controls.Add(row2, 0, 1);
            boxTop.Controls.Add(tlp); mainLayout.Controls.Add(boxTop, 0, 0);
            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true };
            mainLayout.Controls.Add(_dgv, 0, 1);
            return mainLayout;
        }
        private void RefreshGrid() { _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd")); }
    }
}
