using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; // 🟢 必須引用 EPPlus 進行 XLSX 匯出

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

            // --- 第一個框：控制與設定 ---
            GroupBox boxTop = new GroupBox { Text = "數據檢索與欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblDate = new Label { Text = "日期區間:", Location = new Point(20, 40), AutoSize = true };
            _dtpStart = new DateTimePicker { Location = new Point(110, 35), Width = 150, Format = DateTimePickerFormat.Short };
            _dtpStart.Value = DateTime.Now.AddDays(-30); 
            
            Label lblTo = new Label { Text = "至", Location = new Point(270, 40), AutoSize = true };
            _dtpEnd = new DateTimePicker { Location = new Point(300, 35), Width = 150, Format = DateTimePickerFormat.Short };
            
            Button btnRead = new Button { Text = "讀取資料庫", Location = new Point(470, 32), Size = new Size(120, 35), BackColor = Color.LightBlue };
            btnRead.Click += BtnRead_Click;

            Label lblNewCol = new Label { Text = "新增欄位標題:", Location = new Point(20, 90), AutoSize = true };
            _txtNewColName = new TextBox { Location = new Point(150, 85), Width = 200 };
            Button btnAddCol = new Button { Text = "確認新增欄位", Location = new Point(370, 82), Size = new Size(150, 35), BackColor = Color.LightGray };
            btnAddCol.Click += BtnAddCol_Click;

            Button btnImportCsv = new Button { 
                Text = "📥 匯入 CSV", Location = new Point(530, 82), Size = new Size(120, 35), BackColor = Color.Orange 
            };
            btnImportCsv.Click += BtnImportCsv_Click;

            Button btnSaveManual = new Button { 
                Text = "💾 手動儲存所有變更", Location = new Point(20, 140), Size = new Size(340, 40), 
                BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnSaveManual.Click += BtnSaveManual_Click;

            // 🟢 修正 2：刪除按鍵寬度從 150 加寬至 180 (增加 1/5)
            Button btnDeleteRow = new Button {
                Text = "🗑️ 刪除選取資料", Location = new Point(370, 140), Size = new Size(180, 40),
                BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;

            boxTop.Controls.AddRange(new Control[] { lblDate, _dtpStart, lblTo, _dtpEnd, btnRead, lblNewCol, _txtNewColName, btnAddCol, btnImportCsv, btnSaveManual, btnDeleteRow });
            mainLayout.Controls.Add(boxTop, 0, 0);

            // --- 第二個框：數據顯示 ---
            GroupBox boxBottom = new GroupBox { Text = "水處理數據明細 (支援 Ctrl+V 貼上，點擊右鍵匯出)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
            _dgv = new DataGridView {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                SelectionMode = DataGridViewSelectionMode.CellSelect, 
                AllowUserToAddRows = true,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            };
            
            _dgv.KeyDown += Dgv_KeyDown;
            SetupContextMenu();

            boxBottom.Controls.Add(_dgv);
            mainLayout.Controls.Add(boxBottom, 0, 1);

            return mainLayout;
        }

        private void BtnRead_Click(object sender, EventArgs e)
        {
            if (!DataManager.IsDbFileExists())
            {
                if (MessageBox.Show("未找到資料庫，是否新增？", "詢問", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    DataManager.CreateWaterTable();
                else return;
            }
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            DataTable dt = DataManager.GetWaterData(_dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            _dgv.DataSource = dt;
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
        }

        private void BtnSaveManual_Click(object sender, EventArgs e)
        {
            _dgv.EndEdit(); 
            DataTable dt = (DataTable)_dgv.DataSource;

            if (dt == null) return;
            int count = 0;
            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row.RowState == DataRowState.Deleted) continue;
                    DataManager.UpsertWaterRecord(row);
                    count++;
                }
                MessageBox.Show(string.Format("成功處理 {0} 筆資料的新增/複寫作業。", count), "存檔完成");
                RefreshGrid(); 
            }
            catch (Exception ex) { MessageBox.Show("手動存檔失敗：" + ex.Message); }
        }

        private void BtnDeleteRow_Click(object sender, EventArgs e)
        {
            if (_dgv.CurrentRow == null) { MessageBox.Show("請先點選您要刪除的資料行！"); return; }
            if (_dgv.CurrentRow.IsNewRow) { MessageBox.Show("無法刪除尚未存檔的空白行。"); return; }

            var idCell = _dgv.CurrentRow.Cells["Id"].Value;
            if (idCell == null || idCell == DBNull.Value)
            {
                _dgv.Rows.Remove(_dgv.CurrentRow); 
                return;
            }

            if (MessageBox.Show("確定要刪除這筆紀錄嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                DataManager.DeleteWaterRecord(Convert.ToInt32(idCell));
                MessageBox.Show("資料已刪除！");
                RefreshGrid();
            }
        }

        private void BtnAddCol_Click(object sender, EventArgs e)
        {
            string colName = _txtNewColName.Text.Trim();
            if (string.IsNullOrEmpty(colName)) return;
            DataManager.AddColumnToWaterTable(colName);
            MessageBox.Show("欄位新增成功，請重新讀取。");
            _txtNewColName.Clear();
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            if (_dgv.DataSource == null)
            {
                MessageBox.Show("請先點擊「讀取資料庫」初始化表單後，再進行匯入！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV 檔案 (*.csv)|*.csv", Title = "選擇要匯入的 CSV 檔案" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        if (lines.Length <= 1) return;

                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = lines[0].Split(','); 

                        int importCount = 0;
                        for (int i = 1; i < lines.Length; i++)
                        {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            string[] vals = lines[i].Split(',');
                            
                            DataRow newRow = dt.NewRow();
                            for (int h = 0; h < headers.Length; h++)
                            {
                                string colName = headers[h].Trim();
                                if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase)) continue; 
                                
                                if (dt.Columns.Contains(colName) && h < vals.Length)
                                    newRow[colName] = vals[h].Trim();
                            }
                            dt.Rows.Add(newRow);
                            importCount++;
                        }
                        MessageBox.Show($"成功匯入 {importCount} 筆資料到畫面上！\n確認無誤後，請點擊「手動儲存所有變更」寫入資料庫。", "匯入成功");
                    }
                    catch (Exception ex) { MessageBox.Show("匯入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteClipboard();
                e.Handled = true;
            }
        }

        private void PasteClipboard()
        {
            try
            {
                string clipboardText = Clipboard.GetText();
                if (string.IsNullOrEmpty(clipboardText)) return;

                string[] lines = clipboardText.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (_dgv.CurrentCell == null) return;

                int startRow = _dgv.CurrentCell.RowIndex;
                int startCol = _dgv.CurrentCell.ColumnIndex;
                DataTable dt = (DataTable)_dgv.DataSource;

                for (int i = 0; i < lines.Length; i++)
                {
                    string[] cells = lines[i].Split('\t');
                    int currentRow = startRow + i;

                    if (currentRow >= _dgv.Rows.Count || (currentRow == _dgv.Rows.Count - 1 && _dgv.AllowUserToAddRows))
                    {
                        dt.Rows.Add(dt.NewRow());
                    }

                    int c = startCol;
                    foreach (string cell in cells)
                    {
                        if (c < _dgv.Columns.Count)
                        {
                            if (!_dgv.Columns[c].ReadOnly)
                            {
                                _dgv[c, currentRow].Value = cell.TrimEnd('\r');
                            }
                        }
                        c++;
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("貼上失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void SetupContextMenu()
        {
            ContextMenuStrip cms = new ContextMenuStrip();
            cms.Font = new Font("Microsoft JhengHei UI", 12F);
            
            var itemCsv = new ToolStripMenuItem("📤 匯出為 CSV");
            itemCsv.Click += (s, e) => ExportData("csv");
            
            // 🟢 修正 1：改為匯出真正的 XLSX
            var itemExcel = new ToolStripMenuItem("📊 匯出為 Excel (XLSX)");
            itemExcel.Click += (s, e) => ExportData("xlsx");

            cms.Items.Add(itemCsv);
            cms.Items.Add(itemExcel);
            _dgv.ContextMenuStrip = cms;
        }

        private void ExportData(string format)
        {
            if (_dgv.Rows.Count == 0 || _dgv.DataSource == null) return;

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = format == "csv" ? "CSV 檔案 (*.csv)|*.csv" : "Excel 活頁簿 (*.xlsx)|*.xlsx";
                sfd.FileName = "水處理數據_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + "." + format;

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        DataTable dt = (DataTable)_dgv.DataSource;

                        if (format == "csv")
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int i = 0; i < _dgv.Columns.Count; i++)
                                sb.Append(_dgv.Columns[i].HeaderText + ",");
                            sb.AppendLine();

                            foreach (DataGridViewRow row in _dgv.Rows)
                            {
                                if (row.IsNewRow) continue;
                                for (int i = 0; i < _dgv.Columns.Count; i++)
                                {
                                    object val = row.Cells[i].Value;
                                    string cellVal = val != null ? val.ToString().Replace(",", "，").Replace("\n", " ") : ""; 
                                    sb.Append(cellVal + ",");
                                }
                                sb.AppendLine();
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                        }
                        else if (format == "xlsx")
                        {
                            // 🟢 修正 1：透過 EPPlus 建立真正的 XLSX 檔案
                            using (var package = new ExcelPackage())
                            {
                                var ws = package.Workbook.Worksheets.Add("水處理數據");
                                ws.Cells["A1"].LoadFromDataTable(dt, true);
                                ws.Cells.AutoFitColumns(); // 自動調整欄寬
                                package.SaveAs(new FileInfo(sfd.FileName));
                            }
                        }
                        
                        MessageBox.Show("資料匯出成功！", "匯出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }
    }
}
