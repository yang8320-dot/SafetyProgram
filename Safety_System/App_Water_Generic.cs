/// FILE: Safety_System/App_Water_Generic.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_Water_Generic
    {
        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        private Label _lblStartDay, _lblEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        
        // 🟢 將控制按鈕宣告於全域以利狀態控制防連點
        private Button _btnToggle, _btnRead, _btnSave, _btnExport, _btnImport;     

        // 🟢 新增：UI 狀態提示列
        private Label _lblStatus;

        private bool _isFirstLoad = true;
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        private bool _isMonthlyMode = false;
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

        // 🟢 完整保留五大水資源表的 Schema Map
        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            { "WaterMeterReadings", "[日期] TEXT, [星期] TEXT, [用電量] TEXT, [用電量日統計] TEXT, [廢水進流量] TEXT, [廢水進流量日統計] TEXT, [廢水處理量] TEXT, [廢水處理量日統計] TEXT, [水站廢水排放量] TEXT, [水站廢水排放量日統計] TEXT, [納管排放量] TEXT, [納管排放量日統計] TEXT, [回收水6吋] TEXT, [回收水6吋日統計] TEXT, [回收水雙介質A] TEXT, [回收水雙介質A日統計] TEXT, [回收水雙介質B] TEXT, [回收水雙介質B日統計] TEXT, [軟水A通量] TEXT, [軟水B通量] TEXT, [軟水C通量] TEXT, [濃縮水至冷卻水池] TEXT, [濃縮水至冷卻水池日統計] TEXT, [濃縮水至逆洗池] TEXT, [濃縮水至逆洗池日統計] TEXT, [貯存池至循環水池] TEXT, [貯存池至循環水池日統計] TEXT, [製程式至循環水池] TEXT, [製程式至循環水池日統計] TEXT, [污泥產出KG] TEXT, [備註] TEXT" },
            { "WaterChemicals", "[日期] TEXT, [星期] TEXT, [PAC_KG] TEXT, [NAOH_KG] TEXT, [高分子_KG] TEXT, [備註] TEXT" },
            { "WaterUsageDaily", "[日期] TEXT, [星期] TEXT, [廠區自來水使用量] TEXT, [行政區自來水使用量] TEXT, [自來水至貯存池] TEXT, [自來水至貯存池日統計] TEXT, [自來水量至清水池] TEXT, [自來水量至清水池日統計] TEXT, [備註] TEXT" },
            { "DischargeData", "[月份] TEXT, [水量] TEXT, [SS] TEXT, [COD] TEXT, [BOD] TEXT, [氨氮] TEXT, [備註] TEXT" },
            { "WaterVolume", "[月份] TEXT, [廠區自來水繳費單] TEXT, [行政區自來水繳費單] TEXT, [彰濱二廠自來水繳費單] TEXT, [備註] TEXT" }
        };

        public App_Water_Generic(string dbName, string tableName, string chineseTitle)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
        }

        public Control GetView()
        {
            // 初始化資料庫
            string schema = _schemaMap.ContainsKey(_tableName) ? _schemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            // 自動判斷日報/月報模式
            List<string> columns = DataManager.GetColumnNames(_dbName, _tableName);
            if (columns.Contains("月份") && !columns.Contains("日期")) {
                _isMonthlyMode = true;
                _dateColumnName = "月份";
            } else {
                _isMonthlyMode = false;
                _dateColumnName = "日期";
                if (!columns.Contains("星期")) DataManager.AddColumn(_dbName, _tableName, "星期");
            }

            // 🟢 優化 3：TableLayoutPanel 加入 Padding 15，實現完美自適應排版
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, Padding = new Padding(15) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 給 Status Label
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            // 🟢 優化 3：GroupBox 加入 Margin 10px 推開下方距離
            GroupBox boxTop = new GroupBox { Text = $"{_chineseTitle} (庫：{_dbName} 表：{_tableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10), Margin = new Padding(0, 0, 0, 10) };
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

            if (_isMonthlyMode) SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-6));
            else SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            // 初始化全域控制按鈕
            _btnRead = new Button { Text = "🔍 讀取資料", Size = new Size(130, 35), BackColor = Color.WhiteSmoke };
            _btnRead.Click += async (s, e) => { _isFirstLoad = false; await LoadGridDataAsync(); };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(130, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click; 
            
            _btnExport = new Button { Text = "📤 匯出Excel", Size = new Size(130, 35) }; 
            _btnExport.Click += BtnExport_Click;

            _btnImport = new Button { Text = "📥 匯入Excel", Size = new Size(130, 35) }; 
            _btnImport.Click += BtnImportExcel_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(130, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
            };

            _lblStartDay = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            _lblEndDay = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };

            row1.Controls.Add(lblRange);
            row1.Controls.Add(_cboStartYear); row1.Controls.Add(new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboStartMonth); row1.Controls.Add(new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            
            if (!_isMonthlyMode) { row1.Controls.Add(_cboStartDay); row1.Controls.Add(_lblStartDay); }
            
            row1.Controls.Add(new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) });
            row1.Controls.Add(_cboEndYear); row1.Controls.Add(new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            row1.Controls.Add(_cboEndMonth); row1.Controls.Add(new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) });
            
            if (!_isMonthlyMode) { row1.Controls.Add(_cboEndDay); row1.Controls.Add(_lblEndDay); }

            row1.Controls.Add(_btnRead); row1.Controls.Add(_btnExport); row1.Controls.Add(_btnImport); row1.Controls.Add(_btnToggle); row1.Controls.Add(_btnSave);
            boxTop.Controls.Add(row1);

            // ================= 進階管理操作區 =================
            _boxAdvanced = new GroupBox { Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 10) };
            FlowLayoutPanel flpAdv = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            FlowLayoutPanel rowAdv1 = new FlowLayoutPanel { AutoSize = true };
            _txtNewColName = new TextBox { Width = 150 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(100, 35) };
            bAdd.Click += async (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); await LoadGridDataAsync(); _txtNewColName.Clear(); } };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "修改名稱", Size = new Size(100, 35) };
            bRen.Click += async (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); await LoadGridDataAsync(); _txtRenameCol.Clear(); } };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(100, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += async (s, e) => { if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { if(MessageBox.Show($"確定刪除整欄【{_cboColumns.SelectedItem}】？", "確認", MessageBoxButtons.YesNo)==DialogResult.Yes){ DataManager.DropColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString()); await LoadGridDataAsync(); } } };
            
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += async (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>().Select(c => c.OwningRow).Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value).Distinct().ToList();
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    if (AuthManager.VerifyUser()) {
                        foreach (var r in selectedRows) DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                        await LoadGridDataAsync(); MessageBox.Show("刪除成功！");
                    }
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bLimitRead.Click += async (s, e) => { 
                if (int.TryParse(txtLimit.Text, out int l)) { 
                    SetUIState(false, "指定筆數讀取中...", Color.Orange);
                    DataTable dt = null;
                    await Task.Run(() => {
                        dt = DataManager.GetLatestRecords(_dbName, _tableName, l); 
                        EnforceDateFormats(dt);
                    });
                    _dgv.DataSource = dt; 
                    RestoreColumnOrder();
                    SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
                } 
            };
            rowAdv2.Controls.AddRange(new Control[] { new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead });
            
            flpAdv.Controls.Add(rowAdv1); flpAdv.Controls.Add(rowAdv2);
            _boxAdvanced.Controls.Add(flpAdv);

            // 🟢 新增：UI 狀態提示列
            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            // 🟢 優化 3：DGV 行高加大至 35，且加入 Margin 10px 推開下方距離
            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AllowUserToOrderColumns = true,
                Margin = new Padding(0, 10, 0, 10)
            };
            _dgv.RowTemplate.Height = 35;
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _dgv.KeyDown += Dgv_KeyDown;
            
            // 綁定水資源專用自動計算引擎
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            // 🟢 啟動非同步載入
            _ = LoadGridDataAsync(); 
            return main;
        }

        // ==========================================
        // 🟢 狀態與日期強制格式化管理
        // ==========================================
        private void SetUIState(bool isEnabled, string statusText, Color statusColor)
        {
            _btnRead.Enabled = isEnabled;
            _btnSave.Enabled = isEnabled;
            _btnImport.Enabled = isEnabled;
            _btnExport.Enabled = isEnabled;
            
            _lblStatus.Text = statusText;
            _lblStatus.ForeColor = statusColor;
        }

        private void EnforceDateFormats(DataTable dt)
        {
            if (dt == null || !dt.Columns.Contains(_dateColumnName)) return;
            string format = _isMonthlyMode ? "yyyy-MM" : "yyyy-MM-dd";
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                string val = row[_dateColumnName]?.ToString();
                if (!string.IsNullOrWhiteSpace(val)) {
                    val = val.Replace("/", "-");
                    if (DateTime.TryParse(val, out DateTime d)) {
                        row[_dateColumnName] = d.ToString(format);
                    }
                }
            }
        }

        // ==========================================
        // 🟢 核心資料載入 (加入非同步優化)
        // ==========================================
        private async Task LoadGridDataAsync()
        {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            
            string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
            string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);

            await Task.Run(() => {
                if (_isFirstLoad) {
                    dt = DataManager.GetLatestRecords(_dbName, _tableName, 30);
                    _isFirstLoad = false;
                } else {
                    dt = DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate);
                }
                EnforceDateFormats(dt);
            });

            _dgv.DataSource = dt;
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            
            // 強制日期欄位 UI 顯示
            if (_dgv.Columns.Contains(_dateColumnName)) {
                _dgv.Columns[_dateColumnName].DefaultCellStyle.Format = _isMonthlyMode ? "yyyy-MM" : "yyyy-MM-dd";
            }

            UpdateCboColumns();
            RestoreColumnOrder();

            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
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

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            d.SelectedItem = date.Day.ToString("D2");
        }

        private void SaveColumnOrder() { 
            try { 
                var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); 
                File.WriteAllText($"ColOrder_{_dbName}_{_tableName}.txt", string.Join(",", ordered), Encoding.UTF8); 
            } catch { } 
        }
        
        private void RestoreColumnOrder() { 
            try { 
                string fn = $"ColOrder_{_dbName}_{_tableName}.txt"; 
                if (File.Exists(fn)) { 
                    string[] saved = File.ReadAllText(fn, Encoding.UTF8).Split(','); 
                    for (int i = 0; i < saved.Length; i++) 
                        if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; 
                } 
            } catch { } 
        }

        // ==========================================
        // 🟢 核心儲存邏輯 (非同步 + 防連點)
        // ==========================================
        private async void BtnSave_Click(object sender, EventArgs e)
        {
            try {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit();
                SaveColumnOrder();
                
                SetUIState(false, "資料庫寫入中，請稍候...", Color.Orange);

                DataTable dt = (DataTable)_dgv.DataSource;
                bool success = false;
                
                await Task.Run(() => {
                    EnforceDateFormats(dt); 
                    success = DataManager.BulkSaveTable(_dbName, _tableName, dt);
                });

                if (success) {
                    SetUIState(true, "資料儲存成功！", Color.Green);
                    MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    await LoadGridDataAsync();
                } else {
                    SetUIState(true, "資料儲存失敗", Color.Red);
                }
            } 
            catch (Exception ex) {
                SetUIState(true, "儲存異常", Color.Red);
                MessageBox.Show("儲存異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        // ==========================================
        // 🟢 匯入/匯出與剪貼簿 (加入背景運算引擎)
        // ==========================================
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
                        MessageBox.Show("匯出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { 
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
            }
        }

        private async void BtnImportExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "請選擇要匯入的 Excel 檔案" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                        SetUIState(false, "Excel 解析與水資源差值運算中，請稍候...", Color.Orange);

                        DataTable dt = (DataTable)_dgv.DataSource;
                        _dgv.DataSource = null; 
                        
                        await Task.Run(() => {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                if (ws == null || ws.Dimension == null) return;

                                int rowCount = ws.Dimension.Rows;
                                int colCount = ws.Dimension.Columns;

                                string[] headers = new string[colCount];
                                for (int c = 1; c <= colCount; c++) headers[c - 1] = ws.Cells[1, c].Text.Trim();

                                // 🟢 水資源專用：暫停 UI 計算，進入大批運算模式
                                _calcHelper?.BeginBulkUpdate();

                                for (int r = 2; r <= rowCount; r++) {
                                    DataRow nr = dt.NewRow();
                                    bool hasData = false;

                                    for (int c = 1; c <= colCount; c++) {
                                        string cn = headers[c - 1];
                                        string val = ws.Cells[r, c].Text.Trim(); 

                                        if (dt.Columns.Contains(cn) && cn != "Id" && !string.IsNullOrEmpty(val)) {
                                            nr[cn] = val;
                                            hasData = true;
                                        }
                                    }
                                    if (hasData) dt.Rows.Add(nr);
                                }

                                // 🟢 水資源專用：觸發背景中的日差值與星期計算
                                _calcHelper?.RecalculateTable(dt); 
                                _calcHelper?.EndBulkUpdate();
                                EnforceDateFormats(dt); // 匯入後強制格式化日期
                            }
                        });

                        _dgv.DataSource = dt; 
                        RestoreColumnOrder();

                        SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{dt.Rows.Count}", Color.Green);
                        MessageBox.Show("Excel 匯入成功！\n系統已自動計算日差值，請檢查數據後點擊「儲存數據」。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { 
                        await LoadGridDataAsync(); // 發生錯誤時還原資料庫狀態
                        MessageBox.Show("匯入異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    } finally {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

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
                    
                    EnforceDateFormats(dt); // 🟢 貼上後觸發日期格式防呆
                    _dgv.Refresh();
                } catch { _calcHelper?.EndBulkUpdate(); }
            }
        }
    }
}
