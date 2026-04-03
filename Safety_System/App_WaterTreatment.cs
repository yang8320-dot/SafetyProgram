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

        public Control GetView()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 200F)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox boxTop = new GroupBox { Text = "數據檢索與欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            // 🟢 第一行：版面修正為「日期起」與「日期迄」，並強制鎖定格式 yyyy-MM-dd
            Label lblDateStart = new Label { Text = "日期起:", Location = new Point(20, 40), AutoSize = true };
            _dtpStart = new DateTimePicker { 
                Location = new Point(90, 35), Width = 140, 
                Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" 
            };
            _dtpStart.Value = DateTime.Now.AddDays(-30); 
            
            Label lblDateEnd = new Label { Text = "日期迄:", Location = new Point(240, 40), AutoSize = true };
            _dtpEnd = new DateTimePicker { 
                Location = new Point(310, 35), Width = 140, 
                Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" 
            };
            
            Button btnRead = new Button { Text = "讀取資料庫", Location = new Point(470, 32), Size = new Size(120, 35), BackColor = Color.LightBlue };
            btnRead.Click += BtnRead_Click;

            // 第二行
            Label lblNewCol = new Label { Text = "新增欄位標題:", Location = new Point(20, 90), AutoSize = true };
            _txtNewColName = new TextBox { Location = new Point(150, 85), Width = 200 };
            Button btnAddCol = new Button { Text = "確認新增欄位", Location = new Point(370, 82), Size = new Size(150, 35), BackColor = Color.LightGray };
            btnAddCol.Click += BtnAddCol_Click;
            Button btnImportCsv = new Button { Text = "📥 匯入 CSV", Location = new Point(530, 82), Size = new Size(120, 35), BackColor = Color.Orange };
            btnImportCsv.Click += BtnImportCsv_Click;

            // 第三行
            Button btnSaveManual = new Button { 
                Text = "💾 手動儲存所有變更", Location = new Point(20, 140), Size = new Size(340, 40), 
                BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnSaveManual.Click += BtnSaveManual_Click;

            Button btnDeleteRow = new Button {
                Text = "🗑️ 刪除選取資料", Location = new Point(370, 140), Size = new Size(180, 40),
                BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;

            boxTop.Controls.AddRange(new Control[] { lblDateStart, _dtpStart, lblDateEnd, _dtpEnd, btnRead, lblNewCol, _txtNewColName, btnAddCol, btnImportCsv, btnSaveManual, btnDeleteRow });
            mainLayout.Controls.Add(boxTop, 0, 0);

            GroupBox boxBottom = new GroupBox { Text = "數據明細 (支援 Ctrl+V 貼上、右鍵匯出)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
            _dgv = new DataGridView {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AllowUserToAddRows = true
            };
            _dgv.KeyDown += Dgv_KeyDown;
            
            // 手打日期防呆，移開游標立刻轉為 yyyy-MM-dd
            _dgv.CellEndEdit += (s, e) => {
                if (e.RowIndex >= 0 && _dgv.Columns[e.ColumnIndex].Name == "日期") {
                    var cellVal = _dgv[e.ColumnIndex, e.RowIndex].Value;
                    if (cellVal != null && DateTime.TryParse(cellVal.ToString(), out DateTime d)) {
                        _dgv[e.ColumnIndex, e.RowIndex].Value = d.ToString("yyyy-MM-dd");
                    }
                }
            };

            SetupContextMenu();
            boxBottom.Controls.Add(_dgv);
            mainLayout.Controls.Add(boxBottom, 0, 1);

            return mainLayout;
        }

        private void BtnRead_Click(object sender, EventArgs e)
        {
            if (!DataManager.IsDbFileExists()) DataManager.CreateWaterTable();
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            _dgv.DataSource = DataManager.GetWaterData(_dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
        }

        private void BtnSaveManual_Click(object sender, EventArgs e)
        {
            _dgv.EndEdit();
            DataTable dt = (DataTable)_dgv.DataSource;
            if (dt == null) return;
            try {
                foreach (DataRow row in dt.Rows) {
                    if (row.RowState != DataRowState.Deleted) DataManager.UpsertWaterRecord(row);
                }
                MessageBox.Show("存檔完成！");
                RefreshGrid();
            } catch (Exception ex) { MessageBox.Show("存檔失敗：" + ex.Message); }
        }

        private void BtnDeleteRow_Click(object sender, EventArgs e)
        {
            if (_dgv.CurrentRow == null || _dgv.CurrentRow.IsNewRow) return;
            var idVal = _dgv.CurrentRow.Cells["Id"].Value;
            if (idVal != null && idVal != DBNull.Value) {
                if (MessageBox.Show("確定刪除？", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    DataManager.DeleteWaterRecord(Convert.ToInt32(idVal));
                    RefreshGrid();
                }
            } else { _dgv.Rows.Remove(_dgv.CurrentRow); }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V) { PasteClipboard(); e.Handled = true; }
        }

        private void PasteClipboard()
        {
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
                            if (_dgv.Columns[currentCol].Name == "日期" && DateTime.TryParse(val, out DateTime d))
                                val = d.ToString("yyyy-MM-dd");
                            _dgv[currentCol, currentRow].Value = val;
                        }
                    }
                }
            } catch (Exception ex) { MessageBox.Show("貼上失敗：" + ex.Message); }
        }

        private void SetupContextMenu()
        {
            ContextMenuStrip cms = new ContextMenuStrip();
            cms.Items.Add("📤 匯出 CSV", null, (s, e) => ExportData("csv"));
            cms.Items.Add("📊 匯出 XLSX", null, (s, e) => ExportData("xlsx"));
            _dgv.ContextMenuStrip = cms;
        }

        private void ExportData(string format)
        {
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

        private void BtnAddCol_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_txtNewColName.Text)) {
                DataManager.AddColumnToWaterTable(_txtNewColName.Text.Trim());
                MessageBox.Show("新增成功，請重新讀取。");
                _txtNewColName.Clear();
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
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
                                    if (cn == "日期" && DateTime.TryParse(val, out DateTime d)) {
                                        val = d.ToString("yyyy-MM-dd");
                                    }
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
