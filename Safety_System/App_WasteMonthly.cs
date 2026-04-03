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
        private TextBox _txtRenameCol; 
        private ComboBox _cboCols; 

        private const string DbName = "Waste"; 
        private const string TableName = "WasteMonthly"; 

        public Control GetView()
        {
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WasteMonthly] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [代碼] TEXT, [名稱] TEXT, [重量_kg] TEXT, [清理商] TEXT);");

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { Text = "廢棄物月報管理 (庫: Waste / 表: WasteMonthly)", Dock = DockStyle.Fill, Font = new Font("UI", 12F), AutoSize = true };
            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(5) };
            
            _dtpStart = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 160, Format = DateTimePickerFormat.Short };
            Button bRead = new Button { Text = "讀取", Size = new Size(80, 35), BackColor = Color.LightBlue };
            bRead.Click += (s, e) => RefreshGrid();
            
            _cboCols = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtRenameCol = new TextBox { Width = 100 };
            Button bRen = new Button { Text = "✏️ 改名", Size = new Size(80, 35) };
            bRen.Click += (s, e) => { if(!string.IsNullOrEmpty(_cboCols.Text) && VerifyPassword()) { DataManager.RenameColumn(DbName, TableName, _cboCols.Text, _txtRenameCol.Text); RefreshGrid(); } };
            
            Button bSave = new Button { Text = "💾 儲存", Size = new Size(80, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => { _dgv.EndEdit(); DataTable dt = (DataTable)_dgv.DataSource; foreach (DataRow r in dt.Rows) DataManager.UpsertRecord(DbName, TableName, r); MessageBox.Show("完成"); RefreshGrid(); };

            flp.Controls.AddRange(new Control[] { _dtpStart, _dtpEnd, bRead, _cboCols, _txtRenameCol, bRen, bSave });
            box.Controls.Add(flp); main.Controls.Add(box, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            main.Controls.Add(_dgv, 0, 1);
            return main;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            _cboCols.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboCols.Items.Add(c.Name);
        }

        private bool VerifyPassword() {
            Form p = new Form { Width = 450, Height = 240, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 90, Font = new Font("UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 145, Width = 120, Height = 40 };
            p.Controls.AddRange(new Control[] { new Label { Text = "請輸入密碼:", Left = 30, Top = 30, Font = new Font("UI", 14F) }, t, b });
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }
    }
}
