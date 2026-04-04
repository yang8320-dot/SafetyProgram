using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_WaterTreatment
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;

        // 🟢 定義此模組歸屬的資料庫與資料表
        private const string DbName = "Water"; 
        private const string TableName = "WaterMeterReadings"; 

        public Control GetView()
        {
            // 初始化資料表結構
            DataManager.InitTable(DbName, TableName, @"CREATE TABLE IF NOT EXISTS [WaterMeterReadings] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [廢水處理量] TEXT, 
                [廢水進流量] TEXT, 
                [納廢回收6吋] TEXT, 
                [雙介質A] TEXT, 
                [雙介質B] TEXT, 
                [貯存池] TEXT, 
                [軟水A] TEXT, 
                [軟水B] TEXT, 
                [軟水C] TEXT);");

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "廢水處理水量記錄 (庫: Water / 表: WaterMeterReadings)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 25, 10, 10) };
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, AutoSize = true };

            // 第一列：查詢、儲存、匯入與匯出
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _dtpStart = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            _dtpEnd = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short };
            
            Button bRead = new Button { Text = "讀取資料", Size = new Size(100, 35) };
            bRead.Click += (s, e) => RefreshGrid();

            // 🟢 最新防呆存檔邏輯：只要有日期格式錯，就不會寫入資料庫
            Button bSave = new Button { Text = "💾 儲存變更", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            bSave.Click += (s, e) => {
                _dgv.EndEdit();
                DataTable dt = (DataTable)_dgv.DataSource;
                
                // 呼叫底層的全域防呆驗證
                if (DataManager.ValidateAndSaveTable(DbName, TableName, dt)) {
                    MessageBox.Show("儲存完成！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshGrid();
                }
            };

            Button bImport = new Button { Text = "匯入 CSV", Size = new Size(100, 35) };
            bImport.Click += BtnImportCsv_Click;

            Button bExport = new Button { Text = "匯出 Excel", Size = new Size(110, 35) };
            bExport.Click += BtnExportExcel_Click;

            row1.Controls.AddRange(new Control[] { new Label { Text = "區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _dtpStart, _dtpEnd, bRead, bSave, bImport, bExport });

            // 第二列：欄位操作
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
            mainLayout.Controls.Add(boxTop, 0, 0);

            _dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells };
            mainLayout.Controls.Add(_dgv, 0, 1);

            RefreshGrid();
            return mainLayout;
        }

        private void RefreshGrid() {
            _dgv.DataSource = DataManager.GetTableData(DbName, TableName, "日期", _dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) if (c.Name != "Id" && c.Name != "日期") _cboColumns.Items.Add(c.Name);
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV 檔案|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(',');
                        for (int i = 1; i < lines.Length; i++) {
                            DataRow nr = dt.NewRow(); string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) {
                                string cn = headers[h].Trim();
                                if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim();
                            }
                            dt.Rows.Add(nr);
                        }
                    } catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿|*.xlsx", FileName = $"廢水處理水量記錄_{DateTime.Now:yyyyMMdd}" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (var p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("水處理紀錄");
                            DataTable dt = (DataTable)_dgv.DataSource;
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns(); 
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("匯出成功！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private bool VerifyPassword() {
            // 🟢 授權視窗高度優化版 (Height: 270)
            Form p = new Form { Width = 450, Height = 270, Text = "授權驗證", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "請輸入管理員密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox t = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
            Button b = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            p.Controls.Add(lbl); p.Controls.Add(t); p.Controls.Add(b);
            p.AcceptButton = b;
            return p.ShowDialog() == DialogResult.OK && t.Text == "tces";
        }
    }
}
