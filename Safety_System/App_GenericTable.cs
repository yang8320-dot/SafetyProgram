/// FILE: Safety_System/App_GenericTable.cs ///
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
    public class App_GenericTable
    {
        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        private Label _lblStartDay, _lblEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        
        private Button _btnToggle, _btnRead, _btnSave, _btnExport, _btnImport;
        private Label _lblStatus;     

        private bool _isFirstLoad = true;
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        private bool _isMonthlyMode = false;
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            { "ChemRegulations", "[日期] TEXT, [法規來源] TEXT, [公告日期] TEXT, [項目] TEXT, [容許曝露量min] TEXT, [容許曝露量max] TEXT, [管制量min] TEXT, [管制量max] TEXT, [應辦理項目] TEXT, [備註] TEXT" },
            { "SDS_Inventory", "[日期] TEXT, [廠內編號] TEXT, [化學品名稱] TEXT, [CAS_No] TEXT, [危害成份] TEXT, [危害分類] TEXT, [供應商] TEXT, [SDS版本日期] TEXT, [存放地點] TEXT, [最大儲存量] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "AirPollution", "[月份] TEXT, [甲醇] TEXT, [乙醇] TEXT, [油墨] TEXT, [網板清洗劑] TEXT , [備註] TEXT" },
            { "WasteMonthly", "[月份] TEXT, [切_廢玻璃] TEXT, [名稱] TEXT, [重量_kg] TEXT, [清理商] TEXT" },
            { "FireResponsible", "[日期] TEXT, [管轄區域] TEXT, [正負責人] TEXT, [副負責人] TEXT, [聯絡分機] TEXT, [備註] TEXT" },
            { "HazardStats", "[日期] TEXT, [場所名稱] TEXT, [物品名稱] TEXT, [儲存數量] TEXT, [管制倍數] TEXT, [是否合格] TEXT" },
            { "FireEquip", "[日期] TEXT, [設備名稱] TEXT, [編號] TEXT, [位置] TEXT, [有效日期] TEXT, [檢查結果] TEXT, [備註] TEXT" },
            { "訓練時數", "[日期] TEXT, [員工姓名] TEXT, [受訓項目] TEXT, [課程名稱] TEXT, [訓練時數] TEXT, [HR外訓申請] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "EnvMonitor", "[日期] TEXT, [SEG編號] TEXT, [測點名稱] TEXT, [噪音_db] TEXT, [粉塵_區域] TEXT, [粉塵_個人] TEXT, [一氧化鉛] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastewaterPeriodic", "[日期] TEXT, [申報季別] TEXT, [排放水量] TEXT, [COD] TEXT, [SS] TEXT, [BOD] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DrinkingWater", "[日期] TEXT, [採樣點位置] TEXT, [大腸桿菌群] TEXT, [總菌落數] TEXT, [鉛] TEXT, [濁度] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "IndustrialZoneTest", "[日期] TEXT, [採樣點位置] TEXT, [水溫] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [重金屬] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SoilGasTest", "[日期] TEXT, [採樣井編號] TEXT, [測漏氣體濃度] TEXT, [甲烷] TEXT, [二氧化碳] TEXT, [氧氣] TEXT, [檢測機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastewaterSelfTest", "[日期] TEXT, [採樣時間] TEXT, [採樣位置] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [透視度] TEXT, [檢驗人員] TEXT, [備註] TEXT" },
            { "CoolingWaterVendor", "[日期] TEXT, [廠商名稱] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [添加藥劑] TEXT, [檢驗結果] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "CoolingWaterSelf", "[日期] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [檢驗人員] TEXT, [備註] TEXT" },
            { "TCLP", "[日期] TEXT, [樣品名稱] TEXT, [鎘] TEXT, [鉛] TEXT, [鉻] TEXT, [砷] TEXT, [銅] TEXT, [鋅] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterMeterCalibration", "[日期] TEXT, [水錶編號] TEXT, [水錶位置] TEXT, [校正前讀數] TEXT, [校正後讀數] TEXT, [校正單位] TEXT, [下次校正日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "OtherTests", "[日期] TEXT, [檢測項目] TEXT, [檢測位置] TEXT, [檢測數值] TEXT, [單位] TEXT, [合格標準] TEXT, [檢測機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "NearMiss", "[日期] TEXT, [地點] TEXT, [事件經過] TEXT, [提報人] TEXT, [改善措施] TEXT, [附件檔案] TEXT" },
            { "SafetyInspection", "[日期] TEXT, [巡檢區域] TEXT, [檢查項目] TEXT, [檢查結果] TEXT, [缺失描述] TEXT, [改善措施] TEXT, [負責人] TEXT, [狀態] TEXT, [附件檔案] TEXT" },
            { "SafetyObservation", "[日期] TEXT, [區域] TEXT, [類別] TEXT, [描述] TEXT, [觀查人] TEXT, [附件檔案] TEXT" },
            { "TrafficInjury", "[日期] TEXT, [姓名] TEXT, [地點] TEXT, [狀態] TEXT, [附件檔案] TEXT" },
            { "WorkInjury", "[日期] TEXT, [姓名] TEXT, [受傷部位] TEXT, [原因] TEXT, [附件檔案] TEXT" },
            { "HealthPromotion", "[日期] TEXT, [活動名稱] TEXT, [參與人數] TEXT, [執行單位] TEXT, [成果摘要] TEXT, [附件檔案] TEXT, [備註] TEXT" },
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
            string schema = _schemaMap.ContainsKey(_tableName) ? _schemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            List<string> columns = DataManager.GetColumnNames(_dbName, _tableName);
            if (columns.Contains("月份") && !columns.Contains("日期")) {
                _isMonthlyMode = true;
                _dateColumnName = "月份";
            } else {
                _isMonthlyMode = false;
                _dateColumnName = "日期";
            }

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, Padding = new Padding(15) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

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
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(包含所屬的實體附件檔案也將被永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    if (AuthManager.VerifyUser()) {
                        foreach (var r in selectedRows) {
                            // 🟢 同步刪除附件檔案
                            if (_dgv.Columns.Contains("附件檔案")) {
                                string relPath = r.Cells["附件檔案"].Value?.ToString();
                                DeletePhysicalFile(relPath, r.Index);
                            }
                            DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                        }
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
                    SetUIState(false, "讀取中...", Color.Orange);
                    DataTable dt = null;
                    await Task.Run(() => {
                        dt = DataManager.GetLatestRecords(_dbName, _tableName, l); 
                        EnforceDateFormat(dt);
                    });
                    _dgv.DataSource = dt; 
                    ApplyGridStyles();
                    RestoreColumnOrder();
                    SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
                } 
            };
            rowAdv2.Controls.AddRange(new Control[] { new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead });
            
            flpAdv.Controls.Add(rowAdv1); flpAdv.Controls.Add(rowAdv2);
            _boxAdvanced.Controls.Add(flpAdv);

            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AllowUserToOrderColumns = true,
                Margin = new Padding(0, 10, 0, 10)
            };
            _dgv.RowTemplate.Height = 35;
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _dgv.CellFormatting += Dgv_CellFormatting;
            _dgv.CellClick += Dgv_CellClick;
            _dgv.KeyDown += Dgv_KeyDown;
            
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            _ = LoadGridDataAsync();
            return main;
        }

        private void SetUIState(bool isEnabled, string statusText, Color statusColor)
        {
            _btnRead.Enabled = isEnabled;
            _btnSave.Enabled = isEnabled;
            _btnImport.Enabled = isEnabled;
            _btnExport.Enabled = isEnabled;
            
            _lblStatus.Text = statusText;
            _lblStatus.ForeColor = statusColor;
        }

        // ==========================================
        // 🟢 附件檔案專用事件與清理機制
        // ==========================================
        private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                if (colName.Contains("附件檔案") && e.Value != null)
                {
                    string path = e.Value.ToString();
                    if (!string.IsNullOrEmpty(path))
                    {
                        e.Value = Path.GetFileName(path);
                        e.FormattingApplied = true;
                    }
                }
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < _dgv.Rows.Count && !_dgv.Rows[e.RowIndex].IsNewRow)
            {
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                if (colName.Contains("附件檔案"))
                {
                    string currentVal = _dgv[e.ColumnIndex, e.RowIndex].Value?.ToString();

                    using (var frm = new AttachmentForm(currentVal))
                    {
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            if (frm.ResultAction == AttachmentAction.Upload)
                            {
                                try {
                                    string src = frm.SelectedFilePath;
                                    string datePart = DateTime.Now.ToString("yyyy-MM");
                                    
                                    string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件", _dbName, _tableName, datePart);
                                    if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

                                    string ext = Path.GetExtension(src);
                                    string baseName = Path.GetFileNameWithoutExtension(src);
                                    string destName = baseName + ext;
                                    string destPath = Path.Combine(destDir, destName);

                                    int count = 1;
                                    while (File.Exists(destPath)) {
                                        destName = $"{baseName}_{count++}{ext}";
                                        destPath = Path.Combine(destDir, destName);
                                    }

                                    File.Copy(src, destPath);
                                    
                                    string relPath = Path.Combine("附件", _dbName, _tableName, datePart, destName);
                                    _dgv[e.ColumnIndex, e.RowIndex].Value = relPath;
                                    _dgv.EndEdit();
                                } 
                                catch (Exception ex) {
                                    MessageBox.Show("儲存附件失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else if (frm.ResultAction == AttachmentAction.Clear)
                            {
                                // 🟢 呼叫檔案清理，清除實體檔案與空資料夾
                                DeletePhysicalFile(currentVal, e.RowIndex);
                                _dgv[e.ColumnIndex, e.RowIndex].Value = "";
                                _dgv.EndEdit();
                            }
                        }
                    }
                }
            }
        }

        // 🟢 實體檔案清理機制 (包含防呆檢查與資料夾遞迴清理)
        private void DeletePhysicalFile(string relativePath, int currentRowIndex)
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return;

            // 防呆檢查：如果有其他列透過複製貼上指向同一個檔案，則不要刪除實體檔案
            bool isUsedByOthers = false;
            foreach (DataGridViewRow row in _dgv.Rows) {
                if (row.Index == currentRowIndex || row.IsNewRow) continue;
                if (_dgv.Columns.Contains("附件檔案")) {
                    if (row.Cells["附件檔案"].Value?.ToString() == relativePath) {
                        isUsedByOthers = true;
                        break;
                    }
                }
            }

            if (isUsedByOthers) return;

            try {
                string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                if (File.Exists(absPath)) {
                    File.Delete(absPath); // 刪除實體檔案

                    // 檢查並刪除空資料夾，遞迴至「附件」根目錄
                    DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(absPath));
                    string attachRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");

                    while (dir != null && dir.FullName.StartsWith(attachRootDir) && dir.FullName.Length > attachRootDir.Length) {
                        if (dir.Exists && dir.GetFiles().Length == 0 && dir.GetDirectories().Length == 0) {
                            dir.Delete(); // 只有完全空的情況下才會刪除
                            dir = dir.Parent;
                        } else {
                            break; // 若資料夾不為空，就停止遞迴
                        }
                    }
                }
            } 
            catch { /* 安全忽略檔案被鎖定或其他 I/O 異常 */ }
        }

        private void ApplyGridStyles()
        {
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
            
            if (_dgv.Columns.Contains(_dateColumnName)) {
                _dgv.Columns[_dateColumnName].DefaultCellStyle.Format = _isMonthlyMode ? "yyyy-MM" : "yyyy-MM-dd";
            }

            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                if (col.Name.Contains("附件檔案"))
                {
                    col.ReadOnly = true; 
                    col.DefaultCellStyle.ForeColor = Color.Blue;
                    col.DefaultCellStyle.Font = new Font(_dgv.Font, FontStyle.Underline);
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }
        }

        // ==========================================
        // 原有核心邏輯區
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
                    EnforceDateFormat(dt); 
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
                EnforceDateFormat(dt); 
            });

            _dgv.DataSource = dt;
            
            ApplyGridStyles(); 

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

        private void EnforceDateFormat(DataTable dt)
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
                        MessageBox.Show("匯出成功！(附件欄位將輸出為相對路徑，以保證資料完整性)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        SetUIState(false, "Excel 解析與背景運算中，請稍候...", Color.Orange);

                        DataTable dt = (DataTable)_dgv.DataSource;
                        _dgv.DataSource = null; 
                        
                        await Task.Run(() => {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                if (ws == null || ws.Dimension == null) return;

                                int rowCount = ws.Dimension.Rows;
                                int colCount = ws.Dimension.Columns;

                                string[] headers = new string[colCount];
                                for (int c = 1; c <= colCount; c++) {
                                    headers[c - 1] = ws.Cells[1, c].Text.Trim();
                                }

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

                                _calcHelper?.RecalculateTable(dt); 
                                _calcHelper?.EndBulkUpdate();
                                EnforceDateFormat(dt);
                            }
                        });

                        _dgv.DataSource = dt; 
                        ApplyGridStyles(); 
                        RestoreColumnOrder();

                        SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{dt.Rows.Count}", Color.Green);
                        MessageBox.Show("Excel 匯入成功！\n請檢查數據後點擊「儲存數據」。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { 
                        await LoadGridDataAsync(); 
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
                        for (int i = 0; i < cells.Length; i++) {
                            if (c + i < _dgv.Columns.Count) {
                                if (_dgv.Columns[c + i].Name.Contains("附件檔案") || !_dgv.Columns[c + i].ReadOnly) {
                                    _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                                }
                            }
                        }
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt);
                    _calcHelper?.EndBulkUpdate();
                    
                    EnforceDateFormat(dt); 
                    _dgv.Refresh();
                } catch { _calcHelper?.EndBulkUpdate(); }
            }
        }

        // ==========================================
        // 🟢 附件檔案專屬視窗類別
        // ==========================================
        private enum AttachmentAction { None, Upload, Clear }

        private class AttachmentForm : Form
        {
            public string SelectedFilePath { get; private set; }
            public AttachmentAction ResultAction { get; private set; } = AttachmentAction.None;
            
            private string _currentRelPath;
            private string _absPath;

            public AttachmentForm(string currentRelPath)
            {
                _currentRelPath = currentRelPath;
                if (!string.IsNullOrEmpty(_currentRelPath)) {
                    _absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _currentRelPath);
                }

                this.Text = "附件檔案管理";
                this.Size = new Size(450, 350);
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.BackColor = Color.White;

                Label lblStatus = new Label { 
                    Text = string.IsNullOrEmpty(_currentRelPath) ? "狀態: 尚無附件" : "目前附件: " + Path.GetFileName(_currentRelPath), 
                    Dock = DockStyle.Top, 
                    Padding = new Padding(15), 
                    AutoSize = true, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                    ForeColor = string.IsNullOrEmpty(_currentRelPath) ? Color.DimGray : Color.DarkSlateBlue
                };
                this.Controls.Add(lblStatus);

                Button btnOpen = new Button { 
                    Text = "📄 開啟現有附件", 
                    Dock = DockStyle.Top, 
                    Height = 45, 
                    Enabled = !string.IsNullOrEmpty(_currentRelPath) && File.Exists(_absPath), 
                    Font = new Font("Microsoft JhengHei UI", 12F),
                    BackColor = Color.WhiteSmoke
                };
                btnOpen.Click += (s, e) => {
                    try { System.Diagnostics.Process.Start(_absPath); }
                    catch (Exception ex) { MessageBox.Show("無法開啟檔案: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                };
                this.Controls.Add(btnOpen);

                Panel pnlDrop = new Panel { 
                    Dock = DockStyle.Fill, 
                    AllowDrop = true, 
                    BackColor = Color.AliceBlue, 
                    Cursor = Cursors.Hand 
                };
                
                pnlDrop.Paint += (s, e) => {
                    ControlPaint.DrawBorder(e.Graphics, pnlDrop.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Dashed);
                };

                Label lblDrop = new Label { 
                    Text = "📁 點擊此處選擇檔案\n\n或\n\n將檔案拖曳至此區域", 
                    Dock = DockStyle.Fill, 
                    TextAlign = ContentAlignment.MiddleCenter, 
                    Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), 
                    ForeColor = Color.SteelBlue 
                };
                lblDrop.Click += (s, e) => SelectFile();
                pnlDrop.Click += (s, e) => SelectFile();
                pnlDrop.Controls.Add(lblDrop);

                pnlDrop.DragEnter += (s, e) => {
                    if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
                };
                pnlDrop.DragDrop += (s, e) => {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files.Length > 0) {
                        SelectedFilePath = files[0];
                        ResultAction = AttachmentAction.Upload;
                        this.DialogResult = DialogResult.OK;
                    }
                };
                this.Controls.Add(pnlDrop);

                Button btnClear = new Button { 
                    Text = "🗑️ 清除此筆附件", 
                    Dock = DockStyle.Bottom, 
                    Height = 45, 
                    BackColor = Color.IndianRed, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F),
                    Enabled = !string.IsNullOrEmpty(_currentRelPath)
                };
                btnClear.Click += (s, e) => {
                    if (MessageBox.Show("確定要清除此附件記錄嗎？\n(實體檔案將被同步永久刪除)", "確認清除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        ResultAction = AttachmentAction.Clear;
                        this.DialogResult = DialogResult.OK;
                    }
                };
                this.Controls.Add(btnClear);
            }

            private void SelectFile()
            {
                using (OpenFileDialog ofd = new OpenFileDialog { Title = "選擇附件檔案", Filter = "所有檔案 (*.*)|*.*" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK) {
                        SelectedFilePath = ofd.FileName;
                        ResultAction = AttachmentAction.Upload;
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }
    }
}
