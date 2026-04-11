/// FILE: Safety_System/App_GenericTable.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_GenericTable
    {
        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        private Label _lblStartDay, _lblEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        private Button _btnToggle;     

        private bool _isFirstLoad = true;
        
        // 🟢 動態傳入的參數
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        // 🟢 動態時間模式判定
        private bool _isMonthlyMode = false;
        private string _dateColumnName = "日期";

        // 導入自動運算 Helper (處理差值統計與星期推算，並加速大批次貼上)
        private DataGridViewAutoCalcHelper _calcHelper; 

        // 🟢 終極集中管理：所有標準資料表的初始 Schema
        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            // 空污
            //空污申報記錄
            { "AirPollution", "[月份] TEXT, [甲醇] TEXT, [乙醇] TEXT, [油墨] TEXT, [網板清洗劑] TEXT , [備註] TEXT" },
            // 廢棄物
            //廢棄物申報表 生產日月報
            { "WasteMonthly", "[月份] TEXT, [切_廢玻璃] TEXT, [名稱] TEXT, [重量_kg] TEXT, [清理商] TEXT" },
            // 消防
            { "FireResponsible", "[日期] TEXT, [管轄區域] TEXT, [正負責人] TEXT, [副負責人] TEXT, [聯絡分機] TEXT, [備註] TEXT" },
            { "HazardStats", "[日期] TEXT, [場所名稱] TEXT, [物品名稱] TEXT, [儲存數量] TEXT, [管制倍數] TEXT, [是否合格] TEXT" },
            { "FireEquip", "[日期] TEXT, [設備名稱] TEXT, [編號] TEXT, [位置] TEXT, [有效日期] TEXT, [檢查結果] TEXT, [備註] TEXT" },
            // 教育訓練
            { "訓練時數", "[日期] TEXT, [員工姓名] TEXT, [受訓項目] TEXT, [課程名稱] TEXT, [訓練時數] TEXT, [HR外訓申請] TEXT, [備註] TEXT" },
            // 檢測數據
            // 檢測數據 環測
            { "EnvMonitor", "[日期] TEXT, [SEG編號] TEXT, [測點名稱] TEXT, [噪音_db] TEXT, [粉塵_區域] TEXT, [粉塵_個人] TEXT, [一氧化鉛] TEXT, [備註] TEXT" },
            { "WastewaterPeriodic", "[日期] TEXT, [申報季別] TEXT, [排放水量] TEXT, [COD] TEXT, [SS] TEXT, [BOD] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "DrinkingWater", "[日期] TEXT, [採樣點位置] TEXT, [大腸桿菌群] TEXT, [總菌落數] TEXT, [鉛] TEXT, [濁度] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "IndustrialZoneTest", "[日期] TEXT, [採樣點位置] TEXT, [水溫] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [重金屬] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "SoilGasTest", "[日期] TEXT, [採樣井編號] TEXT, [測漏氣體濃度] TEXT, [甲烷] TEXT, [二氧化碳] TEXT, [氧氣] TEXT, [檢測機構] TEXT, [備註] TEXT" },
            { "WastewaterSelfTest", "[日期] TEXT, [採樣時間] TEXT, [採樣位置] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [透視度] TEXT, [檢驗人員] TEXT, [備註] TEXT" },
            { "CoolingWaterVendor", "[日期] TEXT, [廠商名稱] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [添加藥劑] TEXT, [檢驗結果] TEXT, [備註] TEXT" },
            { "CoolingWaterSelf", "[日期] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [檢驗人員] TEXT, [備註] TEXT" },
            { "TCLP", "[日期] TEXT, [樣品名稱] TEXT, [鎘] TEXT, [鉛] TEXT, [鉻] TEXT, [砷] TEXT, [銅] TEXT, [鋅] TEXT, [檢驗機構] TEXT, [備註] TEXT" },
            { "WaterMeterCalibration", "[日期] TEXT, [水錶編號] TEXT, [水錶位置] TEXT, [校正前讀數] TEXT, [校正後讀數] TEXT, [校正單位] TEXT, [下次校正日期] TEXT, [備註] TEXT" },
            { "OtherTests", "[日期] TEXT, [檢測項目] TEXT, [檢測位置] TEXT, [檢測數值] TEXT, [單位] TEXT, [合格標準] TEXT, [檢測機構] TEXT, [備註] TEXT" },
            // 工安
            { "NearMiss", "[日期] TEXT, [地點] TEXT, [事件經過] TEXT, [提報人] TEXT, [改善措施] TEXT" },
            { "SafetyInspection", "[日期] TEXT, [巡檢區域] TEXT, [檢查項目] TEXT, [檢查結果] TEXT, [缺失描述] TEXT, [改善措施] TEXT, [負責人] TEXT, [狀態] TEXT" },
            { "SafetyObservation", "[日期] TEXT, [區域] TEXT, [類別] TEXT, [描述] TEXT, [觀查人] TEXT" },
            { "TrafficInjury", "[日期] TEXT, [姓名] TEXT, [地點] TEXT, [狀態] TEXT" },
            { "WorkInjury", "[日期] TEXT, [姓名] TEXT, [受傷部位] TEXT, [原因] TEXT" },
            // 護理
            // 護理 健康促進
            { "HealthPromotion", "[日期] TEXT, [活動名稱] TEXT, [參與人數] TEXT, [執行單位] TEXT, [成果摘要] TEXT, [備註] TEXT" },
            // 護理 職災申報
            { "WorkInjuryReport", "[月份] TEXT, [男性工時] TEXT, [女性工時] TEXT, [承攬人工時] TEXT, [勞保申請狀態] TEXT, [備註] TEXT" }
        };

        public App_GenericTable(string dbName, string tableName, string chineseTitle)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
        }

        public Control GetView()
        {
            // 1. 初始化資料表
            string schema = _schemaMap.ContainsKey(_tableName) ? _schemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            // 2. 動態偵測查詢模式 (查日期 還是 查月份)
            List<string> columns = DataManager.GetColumnNames(_dbName, _tableName);
            if (columns.Contains("月份") && !columns.Contains("日期")) {
                _isMonthlyMode = true;
                _dateColumnName = "月份";
            } else {
                _isMonthlyMode = false;
                _dateColumnName = "日期";
            }

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"{_chineseTitle} (庫：{_dbName} 表：{_tableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            
            Label lblRange = new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };

            int currentYear = DateTime.Now.Year;
            for (int i = currentYear - 25; i <= currentYear + 25; i++) {
                _cboStartYear.Items.Add(i); _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2")); _cboEndMonth.Items.Add(i.ToString("D2"));
            }
            for (int i = 1; i <= 31; i++) {
                _cboStartDay.Items.Add(i.ToString("D2")); _cboEndDay.Items.Add(i.ToString("D2"));
            }

            // 預設日期設定
            if (_isMonthlyMode) {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-6));
            } else {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            }
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            Button bRead = new Button { Text = "🔍 讀取資料", Size = new Size(150, 35), BackColor = Color.WhiteSmoke };
            bRead.Click += (s, e) => { RefreshGrid(); if (!_isFirstLoad) MessageBox.Show("資料載入完成！"); };

            Button bSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(150, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            bSave.Click += BtnSave_Click; 
            
            Button bExport = new Button { Text = "📤 匯出Excel", Size = new Size(150, 35) }; bExport.Click += BtnExport_Click;
            Button bImport = new Button { Text = "📥 匯入CSV", Size = new Size(150, 35) }; bImport.Click += BtnImportCsv_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
            };

            _lblStartDay = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            _lblEndDay = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };

            row1.Controls.Add(lblRange);
            row1.Controls.Add(_cboStartYear); row1.Controls.Add(new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboStartMonth); row1.Controls.Add(new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            
            // 🟢 動態隱藏日的選項
            if (!_isMonthlyMode) { row1.Controls.Add(_cboStartDay); row1.Controls.Add(_lblStartDay); }
            
            row1.Controls.Add(new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) });
            row1.Controls.Add(_cboEndYear); row1.Controls.Add(new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboEndMonth); row1.Controls.Add(new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            
            // 🟢 動態隱藏日的選項
            if (!_isMonthlyMode) { row1.Controls.Add(_cboEndDay); row1.Controls.Add(_lblEndDay); }

            row1.Controls.Add(bRead); row1.Controls.Add(bExport); row1.Controls.Add(bImport); row1.Controls.Add(_btnToggle); row1.Controls.Add(bSave);
            boxTop.Controls.Add(row1);

            // ==========================================
            // 進階欄位與權限操作區
            // ==========================================
            _boxAdvanced = new GroupBox { Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray };
            FlowLayoutPanel flpAdv = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            FlowLayoutPanel rowAdv1 = new FlowLayoutPanel { AutoSize = true };
            _txtNewColName = new TextBox { Width = 150 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(130, 35) };
            bAdd.Click += (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); RefreshGrid(); _txtNewColName.Clear(); } };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "修改名稱", Size = new Size(130, 35) };
            bRen.Click += (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); RefreshGrid(); _txtRenameCol.Clear(); } };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(130, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += (s, e) => { if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { if(MessageBox.Show($"確定刪除整欄【{_cboColumns.SelectedItem}】？", "確認", MessageBoxButtons.YesNo)==DialogResult.Yes){ DataManager.DropColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString()); RefreshGrid(); } } };
            
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(150, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>().Select(c => c.OwningRow).Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value).Distinct().ToList();
                if (selectedRows.Count > 0) {
                    if (MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                        if (AuthManager.VerifyUser()) {
                            foreach (var r in selectedRows) DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                            RefreshGrid(); MessageBox.Show("刪除成功！");
                        }
                    }
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bLimitRead.Click += (s, e) => { if (int.TryParse(txtLimit.Text, out int l)) { DataTable dt = DataManager.GetLatestRecords(_dbName, _tableName, l); EnforceDateFormat(dt); _dgv.DataSource = dt; RestoreColumnOrder(); } };
            rowAdv2.Controls.AddRange(new Control[] { new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead });
            
            flpAdv.Controls.Add(rowAdv1); flpAdv.Controls.Add(rowAdv2);
            _boxAdvanced.Controls.Add(flpAdv);

            // ==========================================
            // 表格區
            // ==========================================
            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AllowUserToOrderColumns = true 
            };
            
            _dgv.KeyDown += Dgv_KeyDown;
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); main.Controls.Add(_boxAdvanced, 0, 1); main.Controls.Add(_dgv, 0, 2);

            RefreshGrid();
            return main;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit();
                SaveColumnOrder();
                DataTable dt = (DataTable)_dgv.DataSource;
                EnforceDateFormat(dt);
                if (DataManager.BulkSaveTable(_dbName, _tableName, dt)) {
                    MessageBox.Show("儲存完成！");
                    RefreshGrid();
                }
            } finally {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private void RefreshGrid()
        {
            DataTable dt;
            if (_isFirstLoad) {
                dt = DataManager.GetLatestRecords(_dbName, _tableName, 30);
                _isFirstLoad = false;
            } else {
                string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
                string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);
                dt = DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate);
            }
            EnforceDateFormat(dt);
            _dgv.DataSource = dt;
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            UpdateCboColumns();
            RestoreColumnOrder();
        }

        private string GetDateString(ComboBox y, ComboBox m, ComboBox d)
        {
            if (_isMonthlyMode) return $"{y.SelectedItem}-{m.SelectedItem}";
            return $"{y.SelectedItem}-{m.SelectedItem}-{d.SelectedItem}";
        }

        private void UpdateCboColumns()
        {
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns)
                if (c.Name != "Id" && c.Name != _dateColumnName) _cboColumns.Items.Add(c.Name);
        }

        // 🟢 動態時間防呆格式化
        private void EnforceDateFormat(DataTable dt)
        {
            if (dt == null || !dt.Columns.Contains(_dateColumnName)) return;
            string format = _isMonthlyMode ? "yyyy-MM" : "yyyy-MM-dd";
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                string val = row[_dateColumnName]?.ToString();
                if (!string.IsNullOrWhiteSpace(val) && DateTime.TryParse(val, out DateTime d)) 
                    row[_dateColumnName] = d.ToString(format);
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            d.SelectedItem = date.Day.ToString("D2");
        }

        private void SaveColumnOrder() { try { var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); File.WriteAllText($"ColOrder_{_dbName}_{_tableName}.txt", string.Join(",", ordered), Encoding.UTF8); } catch { } }
        private void RestoreColumnOrder() { try { string fn = $"ColOrder_{_dbName}_{_tableName}.txt"; if (File.Exists(fn)) { string[] saved = File.ReadAllText(fn, Encoding.UTF8).Split(','); for (int i = 0; i < saved.Length; i++) if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; } } catch { } }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = _chineseTitle + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) {
                            using (ExcelPackage p = new ExcelPackage()) {
                                var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable(dt, true); p.SaveAs(new FileInfo(sfd.FileName));
                            }
                        } else {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            foreach (DataRow r in dt.Rows) sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功！");
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message); }
                }
            }
        }

        private void BtnImportCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV 檔案 (*.csv)|*.csv" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.Default);
                        if (lines.Length < 2) return;
                        DataTable dt = (DataTable)_dgv.DataSource;
                        string[] headers = ParseCsvLine(lines[0]);

                        _dgv.DataSource = null; 
                        _calcHelper?.BeginBulkUpdate();

                        foreach (string line in lines.Skip(1)) {
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            DataRow nr = dt.NewRow();
                            string[] vs = ParseCsvLine(line);
                            for (int h = 0; h < headers.Length && h < vs.Length; h++) {
                                string cn = headers[h].Trim();
                                if (dt.Columns.Contains(cn) && cn != "Id") nr[cn] = vs[h].Trim().Trim('"');
                            }
                            dt.Rows.Add(nr);
                        }

                        _calcHelper?.RecalculateTable(dt); 
                        _calcHelper?.EndBulkUpdate();
                        _dgv.DataSource = dt; 
                        RestoreColumnOrder();
                        MessageBox.Show("匯入成功!請檢查數據後點擊儲存。");
                    } catch (Exception ex) { RefreshGrid(); MessageBox.Show("匯入異常：" + ex.Message); }
                }
            }
        }

        private string[] ParseCsvLine(string line) { var res = new System.Collections.Generic.List<string>(); bool q = false; var f = new StringBuilder(); foreach (char c in line) { if (c == '\"') q = !q; else if (c == ',' && !q) { res.Add(f.ToString()); f.Clear(); } else f.Append(c); } res.Add(f.ToString()); return res.ToArray(); }

        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V) {
                try {
                    string text = Clipboard.GetText(); if (string.IsNullOrEmpty(text)) return;
                    _calcHelper?.BeginBulkUpdate();
                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex;
                    DataTable dt = (DataTable)_dgv.DataSource;
                    foreach (string line in lines) {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++)
                            if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly)
                                _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt);
                    _calcHelper?.EndBulkUpdate();
                } catch { _calcHelper?.EndBulkUpdate(); }
            }
        }
    }
}
