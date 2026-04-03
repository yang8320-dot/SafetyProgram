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
        private TextBox _txtNewColName;
        
        private ComboBox _cboColumns;
        private TextBox _txtRenameCol;

        public Control GetView()
        {
            // 🟢 1. 最外層主排版：將 Row 0 設為 AutoSize，隨內部控制項自動長高
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { 
                Text = "數據檢索與欄位管理", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F),
                AutoSize = true, // 允許 GroupBox 隨內容擴展
                Padding = new Padding(5, 20, 5, 5)
            };

            // 🟢 2. 控制區內部排版：使用 TableLayoutPanel 分出三行
            TableLayoutPanel tlpControls = new TableLayoutPanel { 
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, AutoSize = true 
            };
            tlpControls.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlpControls.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlpControls.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ==========================================
            // 🟢 第 1 行：日期與讀取、匯入 (使用 FlowLayoutPanel 讓元件自動由左至右流動排列)
            // ==========================================
            FlowLayoutPanel flpRow1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 5, 0, 5) };
            
            Label lblDateStart = new Label { Text = "日期起:", AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            _dtpStart = new DateTimePicker { Width = 170, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            _dtpStart.Value = DateTime.Now.AddDays(-30); 
            
            Label lblDateEnd = new Label { Text = "日期迄:", AutoSize = true, Margin = new Padding(15, 8, 3, 3) };
            _dtpEnd = new DateTimePicker { Width = 170, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            
            Button btnRead = new Button { Text = "讀取資料庫", Size = new Size(120, 35), BackColor = Color.LightBlue, Margin = new Padding(20, 0, 3, 3) };
            btnRead.Click += BtnRead_Click;

            Button btnImportCsv = new Button { Text = "📥 匯入 CSV", Size = new Size(120, 35), BackColor = Color.Orange, Margin = new Padding(10, 0, 3, 3) };
            btnImportCsv.Click += BtnImportCsv_Click;

            flpRow1.Controls.AddRange(new Control[] { lblDateStart, _dtpStart, lblDateEnd, _dtpEnd, btnRead, btnImportCsv });

            // ==========================================
            // 🟢 第 2 行：新增欄位 與 管理欄位
            // ==========================================
            FlowLayoutPanel flpRow2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 5, 0, 5) };
            
            Label lblNewCol = new Label { Text = "新增欄位:", AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            _txtNewColName = new TextBox { Width = 140, Margin = new Padding(3, 5, 3, 3) }; 
            Button btnAddCol = new Button { Text = "確認新增", Size = new Size(100, 35), BackColor = Color.LightGray, Margin = new Padding(10, 0, 3, 3) };
            btnAddCol.Click += BtnAddCol_Click;

            Label lblManageCol = new Label { Text = "管理欄位:", AutoSize = true, Margin = new Padding(30, 8, 3, 3) }; // 增加左距與前方區隔
            _cboColumns = new ComboBox { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3, 5, 3, 3) };
            _txtRenameCol = new TextBox { Width = 130, Margin = new Padding(10, 5, 3, 3) }; 
            
            Button btnRenameCol = new Button { Text = "✏️ 改名", Size = new Size(90, 35), BackColor = Color.LightYellow, Margin = new Padding(10, 0, 3, 3) };
            btnRenameCol.Click += BtnRenameCol_Click;
            
            Button btnDropCol = new Button { Text = "⚠️ 刪除欄位", Size = new Size(120, 35), BackColor = Color.LightPink, Margin = new Padding(10, 0, 3, 3) };
            btnDropCol.Click += BtnDropCol_Click;

            flpRow2.Controls.AddRange(new Control[] { lblNewCol, _txtNewColName, btnAddCol, lblManageCol, _cboColumns, _txtRenameCol, btnRenameCol, btnDropCol });

            // ==========================================
            // 🟢 第 3 行：資料存檔與刪除
            // ==========================================
            FlowLayoutPanel flpRow3 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 5, 0, 10) };
            
            Button btnSaveManual = new Button { 
                Text = "💾 手動儲存所有變更", Size = new Size(340, 40), 
                BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Margin = new Padding(3, 0, 3, 3)
            };
            btnSaveManual.Click += BtnSaveManual_Click;

            Button btnDeleteRow = new Button {
                Text = "🗑️ 刪除選取資料", Size = new Size(180, 40), 
                BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Margin = new Padding(20, 0, 3, 3)
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;

            flpRow3.Controls.AddRange(new Control[] { btnSaveManual, btnDeleteRow });

            // 組合控制區
            tlpControls.Controls.Add(flpRow1, 0, 0);
            tlpControls.Controls.Add(flpRow2, 0, 1);
            tlpControls.Controls.Add(flpRow3, 0, 2);
            boxTop.Controls.Add(tlpControls);
            mainLayout.Controls.Add(boxTop, 0, 0);

            // ==========================================
            // --- 底部明細區 ---
            // ==========================================
            GroupBox boxBottom = new GroupBox { Text = "數據明細 (支援 Ctrl+V 貼上、右鍵匯出)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
            _dgv = new DataGridView {
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, AllowUserToAddRows = true
            };
            _dgv.KeyDown += Dgv_KeyDown;
            
            _dgv.CellEndEdit += (s, e) => {
                if (e.RowIndex >= 0 && _dgv.Columns[e.ColumnIndex].Name == "日期") {
                    var cellVal = _dgv[e.ColumnIndex, e.RowIndex].Value;
                    if (cellVal != null && DateTime.TryParse(cellVal.ToString(), out DateTime d)) _dgv[e.ColumnIndex, e.RowIndex].Value = d.ToString("yyyy-MM-dd");
                }
            };

            SetupContextMenu();
            boxBottom.Controls.Add(_dgv);
            mainLayout.Controls.Add(boxBottom, 0, 1);

            return mainLayout;
        }

        // ==========================================
        // 授權驗證小視窗機制
        // ==========================================
        private bool VerifyPassword()
        {
            Form prompt = new Form() {
                Width = 350, Height = 180, FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "安全授權驗證", StartPosition = FormStartPosition.CenterParent, MaximizeBox = false
            };
            Label textLabel = new Label() { Left = 20, Top = 20, Text = "執行此操作需要管理員權限，\n請輸入授權密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            TextBox textBox = new TextBox() { Left = 20, Top = 70, Width = 280, PasswordChar = '*', Font = new Font("Microsoft JhengHei UI", 12F) };
            Button confirmation = new Button() { Text = "確認執行", Left = 200, Width = 100, Top = 105, DialogResult = DialogResult.OK, Font = new Font("Microsoft JhengHei UI", 10F) };
            
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK && textBox.Text == "tces";
        }

        // ==========================================
        // 重新命名與刪除欄位邏輯
        // ==========================================
        private void BtnRenameCol_Click(object sender, EventArgs e)
        {
            string targetCol = _cboColumns.SelectedItem?.ToString();
            string newName = _txtRenameCol.Text.Trim();
            
            if (string.IsNullOrEmpty(targetCol) || string.IsNullOrEmpty(newName)) {
                MessageBox.Show("請先選擇目標欄位，並在旁邊輸入新的欄位名稱！", "提示"); return;
            }
            if (!VerifyPassword()) { MessageBox.Show("密碼錯誤或取消操作，授權失敗。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            try {
                DataManager.RenameColumnInWaterTable(targetCol, newName);
                MessageBox.Show($"欄位已成功更名為 [{newName}]！", "成功");
                _txtRenameCol.Clear();
                RefreshGrid();
            } catch (Exception ex) { MessageBox.Show("更名失敗：" + ex.Message); }
        }

        private void BtnDropCol_Click(object sender, EventArgs e)
        {
            string targetCol = _cboColumns.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(targetCol)) { MessageBox.Show("請先選擇要刪除的目標欄位！", "提示"); return; }
            
            if (MessageBox.Show($"⚠️ 極度危險操作！\n\n您確定要徹底刪除 [{targetCol}] 欄位及其內部「所有歷史資料」嗎？\n此操作無法復原！", "刪除欄位確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                if (!VerifyPassword()) { MessageBox.Show("密碼錯誤或取消操作，授權失敗。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                try {
                    DataManager.DropColumnFromWaterTable(targetCol);
                    MessageBox.Show($"欄位 [{targetCol}] 已徹底刪除！", "成功");
                    RefreshGrid();
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message); }
            }
        }

        // ==========================================
        // 讀取與綁定邏輯
        // ==========================================
        private void BtnRead_Click(object sender, EventArgs e)
        {
            if (!DataManager.IsDbFileExists()) DataManager.CreateWaterTable();
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            _dgv.DataSource = DataManager.GetWaterData(_dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;

            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (col.Name != "Id" && col.Name != "日期") _cboColumns.Items.Add(col.Name);
            }
            if (_cboColumns.Items.Count > 0) _cboColumns.SelectedIndex = 0;
        }

        private void BtnSaveManual_Click(object sender, EventArgs e) {
            _dgv.EndEdit();
            DataTable dt = (DataTable)_dgv.DataSource;
            if (dt == null) return;
            try {
                foreach (DataRow row in dt.Rows) { if (row.RowState != DataRowState.Deleted) DataManager.UpsertWaterRecord(row); }
                MessageBox.Show("存檔完成！"); RefreshGrid();
            } catch (Exception ex) { MessageBox.Show("存檔失敗：" + ex.Message); }
        }

        private void BtnDeleteRow_Click(object sender, EventArgs e) {
            if (_dgv.CurrentRow == null || _dgv.CurrentRow.IsNewRow) return;
            var idVal = _dgv.CurrentRow.Cells["Id"].Value;
            if (idVal != null && idVal != DBNull.Value) {
                if (MessageBox.Show("確定刪除？", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    DataManager.DeleteWaterRecord(Convert.ToInt32(idVal)); RefreshGrid();
                }
            } else { _dgv.Rows.Remove(_dgv.CurrentRow); }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e) { if (e.Control && e.KeyCode == Keys.V) { PasteClipboard(); e.Handled = true; } }

        private void PasteClipboard() {
            try {
                string text = Clipboard.GetText();
                if (string.IsNullOrEmpty(text)) return;
                if (_dgv.IsCurrentCellInEditMode) _dgv.EndEdit();
                string[] lines = text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                int lineCount = lines.Length;
                if (lineCount > 0 && string.IsNullOrEmpty(lines[lineCount - 1])) lineCount--;
                int startRow = _dgv.CurrentCell.RowIndex;
                int startCol = _dgv.CurrentCell.ColumnIndex;
                DataTable dt = (DataTable)_dgv.DataSource;
                for (int i = 0; i < lineCount; i++) {
                    string[] cells = lines[i].Split('\t');
                    int currentRow = startRow + i;
                    if (currentRow >= _dgv.Rows.Count - (_dgv.AllowUserToAddRows ? 1 : 0)) dt.Rows.Add(dt.NewRow());
                    for (int j = 0; j < cells.Length; j++) {
                        int currentCol = startCol + j;
                        if (currentCol < _dgv.Columns.Count && !_dgv.Columns[currentCol].ReadOnly) {
                            string val = cells[j].Trim();
                            if (_dgv.Columns[currentCol].Name == "日期" && DateTime.TryParse(val, out DateTime d)) val = d.ToString("yyyy-MM-dd");
                            _dgv[currentCol, currentRow].Value = val;
                        }
                    }
                }
            } catch (Exception ex) { MessageBox.Show("貼上失敗：" + ex.Message); }
        }

        private void SetupContextMenu() {
            ContextMenuStrip cms = new ContextMenuStrip();
            cms.Items.Add("📤 匯出 CSV", null, (s, e) => ExportData("csv"));
            cms.Items.Add("📊 匯出 XLSX", null, (s, e) => ExportData("xlsx"));
            _dgv.ContextMenuStrip = cms;
        }

        private void ExportData(string format) {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = format == "csv" ? "CSV|*.csv" : "Excel|*.xlsx" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (format == "csv") {
                            StringBuilder sb = new StringBuilder();
                            foreach (DataColumn col in dt.Columns) sb.Append(col.ColumnName + ",");
                            sb.AppendLine();
                            foreach (DataRow row in dt.Rows) {
                                foreach (var item in row.ItemArray) sb.Append(item.ToString().Replace(",", "，") + ",");
                                sb.AppendLine();
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                        } else {
                            using (var p = new ExcelPackage()) {
                                var ws = p.Workbook.Worksheets.Add("Data");
                                ws.Cells["A1"].LoadFromDataTable(dt, true);
                                ws.Cells.AutoFitColumns();
                                p.SaveAs(new FileInfo(sfd.FileName));
                            }
                        }
                        MessageBox.Show("匯出成功！");
                    } catch (Exception ex) { MessageBox.Show("失敗：" + ex.Message); }
                }
            }
        }

        private void BtnAddCol_Click(object sender, EventArgs e) {
            if (!string.IsNullOrWhiteSpace(_txtNewColName.Text)) {
                DataManager.AddColumnToWaterTable(_txtNewColName.Text.Trim());
                MessageBox.Show("新增成功，請重新讀取。");
                _txtNewColName.Clear();
                RefreshGrid();
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(',');
                        for (int i = 1; i < lines.Length; i++) {
                            DataRow nr = dt.NewRow();
                            string[] vs = lines[i].Split(',');
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) {
                                string cn = headers[h].Trim();
                                if (dt.Columns.Contains(cn) && cn != "Id") {
                                    string val = vs[h].Trim();
                                    if (cn == "日期" && DateTime.TryParse(val, out DateTime d)) val = d.ToString("yyyy-MM-dd");
                                    nr[cn] = val;
                                }
                            }
                            dt.Rows.Add(nr);
                        }
                    } catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message); }
                }
            }
        }
    }
}
