/// FILE: Safety_System/App_GenericTable.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
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
        private enum TimeMode { Date, YearMonth, Year }
        private TimeMode _timeMode = TimeMode.Date;

        // 🟢 搜尋模式列舉，用於保留使用者查詢狀態
        private enum SearchMode { DateRange, Limit, Advanced }
        private SearchMode _currentSearchMode = SearchMode.DateRange;
        private int _currentLimit = 100;

        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        
        private Button _btnToggle, _btnRead, _btnSave, _btnExport, _btnImport, _btnExportPdf, _btnColSettings;
        private Label _lblStatus;     

        private ComboBox _cboSearchColumn;
        private TextBox _txtSearchKeyword;
        private Button _btnAdvancedSearch;

        private bool _isFirstLoad = true;
        private bool _isApplyingWidths = false; 
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();
        private Dictionary<string, int> _columnWidths = new Dictionary<string, int>();

        public App_GenericTable(string dbName, string tableName, string chineseTitle)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
        }

        private string GetExpectedFolderName(string rowDateStr)
        {
            if (string.IsNullOrWhiteSpace(rowDateStr)) return DateTime.Now.ToString("yyyy-MM");
            
            if (_timeMode == TimeMode.Year) 
            {
                if (rowDateStr.Length >= 4) return rowDateStr.Substring(0, 4);
            }
            else 
            {
                if (rowDateStr.Length >= 7) return rowDateStr.Substring(0, 7);
            }
            return DateTime.Now.ToString("yyyy-MM");
        }

        public Control GetView()
        {
            string schema = TableSchemaManager.SchemaMap.ContainsKey(_tableName) 
                            ? TableSchemaManager.SchemaMap[_tableName] 
                            : TableSchemaManager.DefaultCustomSchema;

            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            LoadVisibilitySettings();
            LoadColumnWidths(); 

            List<string> columns = DataManager.GetColumnNames(_dbName, _tableName);
            if (columns.Contains("月份")) 
            {
                try 
                {
                    DataManager.RenameColumn(_dbName, _tableName, "月份", "年月");
                    columns = DataManager.GetColumnNames(_dbName, _tableName);
                } 
                catch { }
            }

            if (columns.Contains("日期")) 
            { 
                _timeMode = TimeMode.Date; 
                _dateColumnName = "日期"; 
            }
            else if (columns.Contains("年月")) 
            { 
                _timeMode = TimeMode.YearMonth; 
                _dateColumnName = "年月"; 
            }
            else if (columns.Contains("年度")) 
            { 
                _timeMode = TimeMode.Year; 
                _dateColumnName = "年度"; 
            }
            else 
            {
                string altDate = columns.FirstOrDefault(c => c.Contains("日期"));
                if (altDate != null) 
                { 
                    _timeMode = TimeMode.Date; 
                    _dateColumnName = altDate; 
                }
                else 
                {
                    altDate = columns.FirstOrDefault(c => c.Contains("年月"));
                    if (altDate != null) 
                    { 
                        _timeMode = TimeMode.YearMonth; 
                        _dateColumnName = altDate; 
                    }
                    else 
                    {
                        altDate = columns.FirstOrDefault(c => c.Contains("年度"));
                        if (altDate != null) 
                        { 
                            _timeMode = TimeMode.Year; 
                            _dateColumnName = altDate; 
                        }
                        else 
                        { 
                            _timeMode = TimeMode.Date; 
                            _dateColumnName = columns.FirstOrDefault(c => c != "Id") ?? "Id"; 
                        }
                    }
                }
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
                _cboStartYear.Items.Add(i); 
                _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2")); 
                _cboEndMonth.Items.Add(i.ToString("D2"));
            }
            for (int i = 1; i <= 31; i++) {
                _cboStartDay.Items.Add(i.ToString("D2")); 
                _cboEndDay.Items.Add(i.ToString("D2"));
            }

            if (_timeMode == TimeMode.YearMonth || _timeMode == TimeMode.Year) 
            {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-6));
            }
            else 
            {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            }
            
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            // 🟢 按鈕事件改為觸發保留狀態的重載
            _btnRead = new Button { Text = "🔍 讀取資料", Size = new Size(130, 35), BackColor = Color.WhiteSmoke };
            _btnRead.Click += async (s, e) => { 
                _isFirstLoad = false; 
                _currentSearchMode = SearchMode.DateRange; 
                await ReloadCurrentDataAsync(); 
            };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(130, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click; 
            
            _btnExport = new Button { Text = "📤 匯出 Excel", Size = new Size(130, 35) }; 
            _btnExport.Click += BtnExport_Click;

            _btnImport = new Button { Text = "📥 匯入 Excel", Size = new Size(130, 35) }; 
            _btnImport.Click += BtnImportExcel_Click;

            _btnExportPdf = new Button { Text = "📄 導出 PDF", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White, Margin = new Padding(0, 0, 10, 0) };
            _btnExportPdf.Click += BtnExportPdf_Click;

            _btnColSettings = new Button { Text = "👁️ 欄位顯示設定", Size = new Size(160, 35), BackColor = Color.LightSlateGray, ForeColor = Color.White, Margin = new Padding(0, 0, 0, 0) };
            _btnColSettings.Click += BtnColSettings_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(160, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
            };

            Label lblSY = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblSM = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblSD = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            Label lblEY = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblEM = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblED = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };

            if (_timeMode == TimeMode.YearMonth) 
            {
                _cboStartDay.Visible = false; lblSD.Visible = false;
                _cboEndDay.Visible = false; lblED.Visible = false;
            } 
            else if (_timeMode == TimeMode.Year) 
            {
                _cboStartDay.Visible = false; lblSD.Visible = false;
                _cboEndDay.Visible = false; lblED.Visible = false;
                _cboStartMonth.Visible = false; lblSM.Visible = false;
                _cboEndMonth.Visible = false; lblEM.Visible = false;
            }

            row1.Controls.AddRange(new Control[] {
                lblRange, _cboStartYear, lblSY, _cboStartMonth, lblSM, _cboStartDay, lblSD,
                lblTilde, _cboEndYear, lblEY, _cboEndMonth, lblEM, _cboEndDay, lblED,
                _btnRead, _btnExport, _btnImport, _btnToggle, _btnSave
            });

            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 10) };
            FlowLayoutPanel flpAdv = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            FlowLayoutPanel rowAdv1 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            
            Label lblAdvOps = new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            _txtNewColName = new TextBox { Width = 130, Margin = new Padding(0, 4, 5, 0) };
            
            // 🟢 新增欄位 (無縫更新)
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(100, 35), Margin = new Padding(0, 0, 15, 0) };
            bAdd.Click += (s, e) => { 
                if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyUser()) 
                { 
                    string newCol = _txtNewColName.Text;
                    DataManager.AddColumn(_dbName, _tableName, newCol); 
                    
                    DataTable dt = (DataTable)_dgv.DataSource;
                    if (!dt.Columns.Contains(newCol)) {
                        dt.Columns.Add(newCol, typeof(string));
                    }
                    ApplyGridStyles();
                    UpdateCboColumns();
                    _txtNewColName.Clear(); 
                    MessageBox.Show("欄位新增成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } 
            };
            
            _cboColumns = new ComboBox { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 5, 0) }; 
            _txtRenameCol = new TextBox { Width = 110, Margin = new Padding(0, 4, 5, 0) };
            
            // 🟢 修改欄位名稱 (無縫更新)
            Button bRen = new Button { Text = "修改名稱", Size = new Size(100, 35), Margin = new Padding(0, 0, 5, 0) };
            bRen.Click += (s, e) => { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyUser()) 
                { 
                    string oldName = _cboColumns.SelectedItem.ToString();
                    string newName = _txtRenameCol.Text;
                    DataManager.RenameColumn(_dbName, _tableName, oldName, newName); 
                    
                    DataTable dt = (DataTable)_dgv.DataSource;
                    if (dt.Columns.Contains(oldName)) dt.Columns[oldName].ColumnName = newName;
                    if (_dgv.Columns.Contains(oldName)) {
                        _dgv.Columns[oldName].HeaderText = newName;
                        _dgv.Columns[oldName].Name = newName;
                    }
                    UpdateCboColumns();
                    _txtRenameCol.Clear(); 
                    MessageBox.Show("欄位名稱修改成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } 
            };
            
            // 🟢 刪除整欄 (無縫更新)
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(100, 35), BackColor = Color.DarkOrange, ForeColor = Color.White, Margin = new Padding(0, 0, 15, 0) };
            bDelCol.Click += (s, e) => { 
                if (_cboColumns.SelectedItem != null && AuthManager.VerifyUser()) 
                { 
                    string colToDrop = _cboColumns.SelectedItem.ToString();
                    if(MessageBox.Show($"確定刪除整欄【{colToDrop}】？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    { 
                        DataManager.DropColumn(_dbName, _tableName, colToDrop); 
                        
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (dt.Columns.Contains(colToDrop)) dt.Columns.Remove(colToDrop);
                        if (_dgv.Columns.Contains(colToDrop)) _dgv.Columns.Remove(colToDrop);
                        
                        UpdateCboColumns();
                        MessageBox.Show("欄位刪除成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                } 
            };
            
            // 🟢 刪除選取列 (無縫更新)
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(130, 35), BackColor = Color.IndianRed, ForeColor = Color.White, Margin = new Padding(0, 0, 15, 0) };
            bDelRow.Click += (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>().Select(c => c.OwningRow).Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value).Distinct().ToList();
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(包含所屬的實體附件檔案也將被永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                {
                    if (AuthManager.VerifyUser()) 
                    {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        foreach (var r in selectedRows) 
                        {
                            if (_dgv.Columns.Contains("附件檔案")) 
                            {
                                string relPathStr = r.Cells["附件檔案"].Value?.ToString();
                                if (!string.IsNullOrEmpty(relPathStr)) 
                                {
                                    string[] paths = relPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var p in paths) DeletePhysicalFile(p, r.Index);
                                }
                            }
                            int id = Convert.ToInt32(r.Cells["Id"].Value);
                            DataManager.DeleteRecord(_dbName, _tableName, id);
                            
                            // 🟢 修正 CS0411 錯誤：使用 Cast<DataRow>() 替代 AsEnumerable()
                            DataRow rowToDelete = dt.Rows.Cast<DataRow>().FirstOrDefault(dr => dr.RowState != DataRowState.Deleted && Convert.ToInt32(dr["Id"]) == id);
                            if (rowToDelete != null) rowToDelete.Delete();
                        }
                        dt.AcceptChanges(); // 套用變更，讓畫面立刻刷新而不需要查 DB
                        MessageBox.Show("刪除成功！");
                    }
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { lblAdvOps, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow, _btnExportPdf, _btnColSettings });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 5, 0, 0), WrapContents = false };
            
            Label lblLimit = new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            TextBox txtLimit = new TextBox { Width = 70, Text = "100", Margin = new Padding(0, 4, 5, 0) };
            Button bLimitRead = new Button { Text = "讀取", Size = new Size(80, 35), BackColor = Color.SteelBlue, ForeColor = Color.White, Margin = new Padding(0, 0, 20, 0) };
            bLimitRead.Click += async (s, e) => { 
                if (int.TryParse(txtLimit.Text, out int l)) 
                { 
                    _isFirstLoad = false;
                    _currentSearchMode = SearchMode.Limit;
                    _currentLimit = l;
                    await ReloadCurrentDataAsync();
                } 
            };

            Label lblSearchData = new Label { Text = "查詢資料:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            _cboSearchColumn = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 5, 0) };
            
            Label lblKeyword = new Label { Text = "關鍵字(包含):", AutoSize = true, Margin = new Padding(15, 8, 5, 0) };
            _txtSearchKeyword = new TextBox { Width = 180, Margin = new Padding(0, 4, 5, 0) };
            
            _btnAdvancedSearch = new Button { Text = "🔍 條件搜尋", Size = new Size(130, 35), BackColor = Color.SteelBlue, ForeColor = Color.White, Margin = new Padding(0, 0, 0, 0) };
            _btnAdvancedSearch.Click += async (s, e) => {
                _isFirstLoad = false;
                _currentSearchMode = SearchMode.Advanced;
                await ReloadCurrentDataAsync();
            };
            
            rowAdv2.Controls.AddRange(new Control[] { lblLimit, txtLimit, bLimitRead, lblSearchData, _cboSearchColumn, lblKeyword, _txtSearchKeyword, _btnAdvancedSearch });

            flpAdv.Controls.Add(rowAdv1); 
            flpAdv.Controls.Add(rowAdv2);
            _boxAdvanced.Controls.Add(flpAdv);

            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = true, 
                AllowUserToResizeColumns = true, 
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, 
                AllowUserToOrderColumns = true,
                Margin = new Padding(0, 10, 0, 10)
            };
            
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _dgv.CellFormatting += Dgv_CellFormatting;
            _dgv.CellClick += Dgv_CellClick;
            _dgv.KeyDown += Dgv_KeyDown;
            
            _dgv.KeyPress += Dgv_KeyPress;
            _dgv.EditingControlShowing += Dgv_EditingControlShowing;

            _dgv.DataError += (s, e) => { e.ThrowException = false; };
            _dgv.CurrentCellDirtyStateChanged += Dgv_CurrentCellDirtyStateChanged;
            _dgv.CellValueChanged += Dgv_CellValueChanged;
            
            _dgv.ColumnWidthChanged += Dgv_ColumnWidthChanged;

            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            _ = ReloadCurrentDataAsync(); 
            return main;
        }

        // ==========================================
        // 🟢 狀態保持與重載機制 ( ReloadCurrentDataAsync )
        // ==========================================
        private async Task ReloadCurrentDataAsync()
        {
            // 1. 記錄目前的捲動位置與選取狀態
            int firstRowIndex = -1;
            int selectedRowIndex = -1;
            int selectedColIndex = -1;
            
            if (_dgv.Rows.Count > 0 && !_isFirstLoad) {
                firstRowIndex = _dgv.FirstDisplayedScrollingRowIndex;
                if (_dgv.CurrentCell != null) {
                    selectedRowIndex = _dgv.CurrentCell.RowIndex;
                    selectedColIndex = _dgv.CurrentCell.ColumnIndex;
                }
            }

            // 2. 根據目前的查詢模式，重載資料庫
            if (_currentSearchMode == SearchMode.Advanced) {
                await ExecuteAdvancedSearchAsync();
            } else if (_currentSearchMode == SearchMode.Limit) {
                await LoadLimitDataAsync(_currentLimit);
            } else {
                await LoadGridDataAsync();
            }

            // 3. 恢復捲動位置與選取狀態
            try {
                if (firstRowIndex >= 0 && firstRowIndex < _dgv.Rows.Count)
                    _dgv.FirstDisplayedScrollingRowIndex = firstRowIndex;
                
                if (selectedRowIndex >= 0 && selectedRowIndex < _dgv.Rows.Count && selectedColIndex >= 0) {
                    _dgv.ClearSelection();
                    _dgv.CurrentCell = _dgv.Rows[selectedRowIndex].Cells[selectedColIndex];
                    _dgv.Rows[selectedRowIndex].Selected = true;
                }
            } catch { } // 避免因為資料量變少導致索引越界
        }

        private async Task LoadGridDataAsync() 
        {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            
            string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
            string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);

            await Task.Run(() => {
                if (_isFirstLoad) 
                {
                    dt = DataManager.GetLatestRecords(_dbName, _tableName, 30);
                    _isFirstLoad = false;
                } 
                else 
                {
                    dt = DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate);
                }
                EnforceDateFormats(dt);
            });

            _dgv.DataSource = dt;
            ApplyGridStyles();
            UpdateCboColumns();
            RestoreColumnOrder();

            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
        }

        private async Task LoadLimitDataAsync(int limit) 
        {
            SetUIState(false, "讀取中...", Color.Orange);
            DataTable dt = null;
            await Task.Run(() => { 
                dt = DataManager.GetLatestRecords(_dbName, _tableName, limit); 
                EnforceDateFormats(dt); 
            });
            _dgv.DataSource = dt; 
            ApplyGridStyles(); 
            UpdateCboColumns();
            RestoreColumnOrder();
            SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
        }

        private async Task ExecuteAdvancedSearchAsync()
        {
            SetUIState(false, "條件搜尋中，請稍候...", Color.Orange);
            
            string searchCol = _cboSearchColumn.SelectedItem?.ToString();
            string keyword = _txtSearchKeyword.Text;

            DataTable resultDt = null;

            await Task.Run(() => {
                DataTable allData = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                DataView dv = allData.DefaultView;

                if (!string.IsNullOrEmpty(searchCol)) 
                {
                    // 🟢 智慧空值搜尋邏輯
                    if (string.IsNullOrWhiteSpace(keyword)) {
                        dv.RowFilter = $"[{searchCol}] IS NULL OR [{searchCol}] = ''";
                    } else {
                        dv.RowFilter = $"[{searchCol}] LIKE '%{keyword.Replace("'", "''")}%'";
                    }
                }
                
                dv.Sort = "Id DESC"; 
                
                resultDt = dv.ToTable(); 
                EnforceDateFormats(resultDt);
            });

            _dgv.DataSource = resultDt;
            ApplyGridStyles();
            UpdateCboColumns();
            RestoreColumnOrder();
            SetUIState(true, $"搜尋完成，共找到 {resultDt.Rows.Count} 筆符合條件資料", Color.Green);
        }

        // ==========================================
        // 欄位寬度與排序管理
        // ==========================================
        private void LoadColumnWidths()
        {
            _columnWidths.Clear();
            var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Width");
            foreach (var kvp in dict) {
                if (int.TryParse(kvp.Value, out int w)) _columnWidths[kvp.Key] = w;
            }
        }

        private void SaveColumnWidths()
        {
            foreach (var kvp in _columnWidths) {
                DataManager.SaveGridConfig(_dbName, _tableName, "Width", kvp.Key, kvp.Value.ToString());
            }
        }

        private void Dgv_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            if (_isFirstLoad || _isApplyingWidths) return;
            if (e.Column != null) {
                _columnWidths[e.Column.Name] = e.Column.Width;
                SaveColumnWidths();
            }
        }

        private void LoadVisibilitySettings()
        {
            _columnVisibility.Clear();
            var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Visibility");
            foreach (var kvp in dict) {
                _columnVisibility[kvp.Key] = (kvp.Value == "1");
            }
        }

        private void SaveVisibilitySettings()
        {
            DataManager.ClearGridConfig(_dbName, _tableName, "Visibility");
            foreach (var kvp in _columnVisibility) {
                DataManager.SaveGridConfig(_dbName, _tableName, "Visibility", kvp.Key, kvp.Value ? "1" : "0");
            }
        }

        private void SaveColumnOrder() 
        { 
            try { 
                var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); 
                DataManager.SaveGridConfig(_dbName, _tableName, "Order", "All", string.Join(",", ordered));
            } catch { } 
        }
        
        private void RestoreColumnOrder() 
        { 
            try { 
                var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Order");
                if (dict.ContainsKey("All")) { 
                    string[] saved = dict["All"].Split(','); 
                    for (int i = 0; i < saved.Length; i++) {
                        if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; 
                    }
                } 
            } catch { } 
        }

        private void BtnColSettings_Click(object sender, EventArgs e)
        {
            if (_dgv.Columns.Count == 0) return;

            using (Form f = new Form { Text = "👁️ 欄位顯示設定", Size = new Size(350, 500), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false }) {
                
                Label lblTop = new Label { Text = "請勾選欲顯示在表格中的欄位：", Dock = DockStyle.Top, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.SteelBlue };
                f.Controls.Add(lblTop);

                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.None, Padding = new Padding(10) };
                
                foreach (DataGridViewColumn col in _dgv.Columns) {
                    if (col.Name == "Id") continue;
                    bool isChecked = _columnVisibility.ContainsKey(col.Name) ? _columnVisibility[col.Name] : true;
                    clbCols.Items.Add(col.Name, isChecked);
                }

                f.Controls.Add(clbCols);

                Button btnSave = new Button { Text = "💾 儲存並套用設定", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnSave.Click += (s, ev) => {
                    for (int i = 0; i < clbCols.Items.Count; i++) {
                        string colName = clbCols.Items[i].ToString();
                        bool isChecked = clbCols.GetItemChecked(i);
                        _columnVisibility[colName] = isChecked;
                        if (_dgv.Columns.Contains(colName)) {
                            _dgv.Columns[colName].Visible = isChecked;
                        }
                    }
                    SaveVisibilitySettings();
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(btnSave);
                f.ShowDialog();
            }
        }

        private void SetUIState(bool isEnabled, string statusText, Color statusColor) 
        {
            _btnRead.Enabled = isEnabled; 
            _btnSave.Enabled = isEnabled; 
            _btnImport.Enabled = isEnabled; 
            _btnExport.Enabled = isEnabled;
            _btnExportPdf.Enabled = isEnabled;
            _btnColSettings.Enabled = isEnabled;
            _btnAdvancedSearch.Enabled = isEnabled;
            
            _lblStatus.Text = statusText; 
            _lblStatus.ForeColor = statusColor;
        }

        private void Dgv_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (_dgv.CurrentCell != null && !_dgv.CurrentCell.ReadOnly && !_dgv.IsCurrentCellInEditMode)
            {
                if (char.IsLetterOrDigit(e.KeyChar) || char.IsPunctuation(e.KeyChar) || char.IsSymbol(e.KeyChar) || char.IsWhiteSpace(e.KeyChar))
                {
                    _dgv.BeginEdit(true);
                    if (_dgv.EditingControl is TextBox txt)
                    {
                        txt.Text = e.KeyChar.ToString();
                        txt.SelectionStart = txt.Text.Length;
                        e.Handled = true;
                    }
                }
            }
        }

        private void Dgv_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox cbo)
            {
                cbo.DropDownStyle = ComboBoxStyle.DropDownList;

                if (_tableName == "SafetyInspection" && _dgv.CurrentCell != null)
                {
                    string colName = _dgv.Columns[_dgv.CurrentCell.ColumnIndex].Name;
                    if (colName == "危害類型細分類")
                    {
                        string parentVal = _dgv.CurrentRow.Cells["危害類型主項"].Value?.ToString() ?? "";
                        var items = TableSchemaManager.GetDependentDropdownList(_tableName, colName, parentVal);
                        object currentVal = _dgv.CurrentCell.Value;
                        cbo.Items.Clear(); cbo.Items.AddRange(items);
                        if (currentVal != null && cbo.Items.Contains(currentVal)) cbo.SelectedItem = currentVal;
                    }
                    else if (colName == "違規樣態類型")
                    {
                        string parentVal = _dgv.CurrentRow.Cells["危害類型細分類"].Value?.ToString() ?? "";
                        var items = TableSchemaManager.GetDependentDropdownList(_tableName, colName, parentVal);
                        object currentVal = _dgv.CurrentCell.Value;
                        cbo.Items.Clear(); cbo.Items.AddRange(items);
                        if (currentVal != null && cbo.Items.Contains(currentVal)) cbo.SelectedItem = currentVal;
                    }
                }
            }
            else if (e.Control is TextBox txt)
            {
                txt.Multiline = true;
                txt.KeyDown -= TextBox_KeyDown;
                txt.KeyDown += TextBox_KeyDown;
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Alt && e.KeyCode == Keys.Enter)
            {
                if (sender is TextBox txt)
                {
                    int selectionStart = txt.SelectionStart;
                    txt.Text = txt.Text.Insert(selectionStart, Environment.NewLine);
                    txt.SelectionStart = selectionStart + Environment.NewLine.Length;
                    e.Handled = true;
                }
            }
        }

        private void BtnExportPdf_Click(object sender, EventArgs e)
        {
            if (_dgv.Rows.Count <= 1) {
                MessageBox.Show("目前沒有資料可供導出。");
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = $"{_chineseTitle}_{DateTime.Now:yyyyMMdd}" }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    Form activeForm = Form.ActiveForm;
                    if (activeForm != null) activeForm.Cursor = Cursors.WaitCursor;

                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = true; 
                    pd.DefaultPageSettings.Margins = new Margins(30, 30, 40, 40);
                    
                    int rowIndex = 0;
                    int pageNumber = 1;

                    int rowsPerPageEstimate = 20; 
                    int totalPages = (int)Math.Ceiling((double)(_dgv.Rows.Count - 1) / rowsPerPageEstimate);

                    pd.PrintPage += (s, ev) => {
                        Graphics g = ev.Graphics;
                        float x = ev.MarginBounds.Left;
                        float y = ev.MarginBounds.Top;
                        float pageWidth = ev.MarginBounds.Width;

                        Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                        Font fSubTitle = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold); 
                        Font fBody = new Font("Microsoft JhengHei UI", 9F);
                        Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

                        StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                        g.DrawString("台灣玻璃工業股份有限公司-彰濱廠", fTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 40), sfCenter); y += 35;
                        g.DrawString(_chineseTitle, fSubTitle, Brushes.Black, new RectangleF(x, y, pageWidth, 30), sfCenter); y += 30;
                        
                        string filterStr = "";
                        if (!string.IsNullOrEmpty(_txtSearchKeyword.Text)) filterStr = $" | 關鍵字: {_txtSearchKeyword.Text}";
                        g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}{filterStr}", fBody, Brushes.Gray, new RectangleF(x, y, pageWidth, 25), sfLeft); y += 25;

                        var visCols = _dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList();
                        if (visCols.Count == 0) return;

                        float totalGridWidth = visCols.Sum(c => c.Width);
                        float[] actualColWidths = new float[visCols.Count];
                        
                        for (int i = 0; i < visCols.Count; i++) {
                            actualColWidths[i] = (visCols[i].Width / totalGridWidth) * pageWidth;
                        }

                        float currX = x;
                        float rowH = 32;

                        for (int i = 0; i < visCols.Count; i++)
                        {
                            RectangleF rect = new RectangleF(currX, y, actualColWidths[i], rowH);
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            g.DrawString(visCols[i].HeaderText, fHead, Brushes.Black, rect, sfCenter);
                            currX += actualColWidths[i];
                        }
                        y += rowH;

                        while (rowIndex < _dgv.Rows.Count)
                        {
                            if (_dgv.Rows[rowIndex].IsNewRow) { rowIndex++; continue; }

                            float maxRowH = rowH;
                            for (int i = 0; i < visCols.Count; i++) {
                                string val = _dgv[visCols[i].Index, rowIndex].Value?.ToString() ?? "";
                                SizeF sSize = g.MeasureString(val, fBody, (int)actualColWidths[i], sfLeft);
                                if (sSize.Height + 10 > maxRowH) maxRowH = sSize.Height + 10;
                            }

                            if (y + maxRowH > ev.MarginBounds.Bottom - 30) 
                            {
                                g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fBody, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);
                                pageNumber++;
                                ev.HasMorePages = true;
                                return;
                            }

                            currX = x;
                            for (int i = 0; i < visCols.Count; i++)
                            {
                                RectangleF rect = new RectangleF(currX, y, actualColWidths[i], maxRowH);
                                g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                                string val = _dgv[visCols[i].Index, rowIndex].Value?.ToString() ?? "";
                                
                                RectangleF textRect = new RectangleF(rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4);
                                g.DrawString(val, fBody, Brushes.Black, textRect, sfLeft);
                                currX += actualColWidths[i];
                            }
                            y += maxRowH;
                            rowIndex++;
                        }
                        
                        g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fBody, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);
                        ev.HasMorePages = false;
                        rowIndex = 0; 
                        pageNumber = 1;
                    };

                    try {
                        pd.Print();
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                        MessageBox.Show("PDF 報表匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void Dgv_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (_dgv.IsCurrentCellDirty && _dgv.CurrentCell is DataGridViewComboBoxCell)
            {
                _dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            
            string colName = _dgv.Columns[e.ColumnIndex].Name;

            if (_tableName == "SafetyInspection")
            {
                if (colName == "危害類型主項")
                {
                    if (_dgv.Columns.Contains("危害類型細分類")) _dgv.Rows[e.RowIndex].Cells["危害類型細分類"].Value = "";
                    if (_dgv.Columns.Contains("違規樣態類型")) _dgv.Rows[e.RowIndex].Cells["違規樣態類型"].Value = "";
                }
                else if (colName == "危害類型細分類")
                {
                    if (_dgv.Columns.Contains("違規樣態類型")) _dgv.Rows[e.RowIndex].Cells["違規樣態類型"].Value = "";
                }
            }
        }

        private void ApplyGridStyles() 
        {
            _isApplyingWidths = true; 

            if (_dgv.Columns.Contains("Id")) 
            {
                _dgv.Columns["Id"].ReadOnly = true;
                _dgv.Columns["Id"].Visible = false;
            }
            
            if (_dgv.Columns.Contains(_dateColumnName)) 
            {
                string fmt = "yyyy-MM-dd";
                if (_timeMode == TimeMode.YearMonth) fmt = "yyyy-MM";
                else if (_timeMode == TimeMode.Year) fmt = "yyyy";
                
                _dgv.Columns[_dateColumnName].DefaultCellStyle.Format = fmt;
            }
            
            foreach (DataGridViewColumn col in _dgv.Columns) 
            {
                if (_columnVisibility.ContainsKey(col.Name)) {
                    col.Visible = _columnVisibility[col.Name];
                }

                if (col.Name.Contains("附件檔案")) 
                {
                    col.ReadOnly = true; 
                    col.DefaultCellStyle.ForeColor = Color.Blue;
                    col.DefaultCellStyle.Font = new Font(_dgv.Font, FontStyle.Underline);
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else 
                {
                    col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }

            SetupDropdownColumns();
            
            _dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            
            _dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                if (_columnWidths.ContainsKey(col.Name) && _columnWidths[col.Name] > 0)
                {
                    col.Width = _columnWidths[col.Name];
                }
            }

            _isApplyingWidths = false; 
        }

        private void SetupDropdownColumns()
        {
            foreach (DataGridViewColumn col in _dgv.Columns.Cast<DataGridViewColumn>().ToList())
            {
                string[] items = TableSchemaManager.GetDropdownList(_tableName, col.Name);
                if (items != null && !(_dgv.Columns[col.Name] is DataGridViewComboBoxColumn))
                {
                    int colIndex = col.Index;
                    _dgv.Columns.RemoveAt(colIndex);

                    DataGridViewComboBoxColumn cboCol = new DataGridViewComboBoxColumn
                    {
                        Name = col.Name,
                        HeaderText = col.HeaderText,
                        DataPropertyName = col.DataPropertyName,
                        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox,
                        FlatStyle = FlatStyle.Flat,
                        SortMode = DataGridViewColumnSortMode.Automatic
                    };

                    if (_columnVisibility.ContainsKey(col.Name)) cboCol.Visible = _columnVisibility[col.Name];

                    List<string> finalItems = new List<string>(items);
                    if (_dgv.DataSource is DataTable dt)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            string val = row[col.Name]?.ToString().Trim();
                            if (!string.IsNullOrEmpty(val) && !finalItems.Contains(val))
                            {
                                finalItems.Add(val);
                            }
                        }
                    }
                    cboCol.Items.AddRange(finalItems.ToArray());
                    _dgv.Columns.Insert(colIndex, cboCol);
                }
            }
        }

        private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) 
            {
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                if (colName.Contains("附件檔案") && e.Value != null) 
                {
                    string pathStr = e.Value.ToString();
                    if (!string.IsNullOrEmpty(pathStr)) 
                    {
                        string[] parts = pathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 1) 
                        {
                            e.Value = $"📁 [共 {parts.Length} 個檔案]";
                        } 
                        else 
                        {
                            e.Value = Path.GetFileName(parts[0]);
                        }
                        e.FormattingApplied = true;
                    }
                }
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e) 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < _dgv.Rows.Count && !_dgv.Rows[e.RowIndex].IsNewRow) 
            {
                if (_dgv.Columns[e.ColumnIndex].Name.Contains("附件檔案")) 
                {
                    string currentVal = _dgv[e.ColumnIndex, e.RowIndex].Value?.ToString();
                    
                    string rowDateStr = _dgv[_dateColumnName, e.RowIndex].Value?.ToString() ?? "";
                    string targetFolder = GetExpectedFolderName(rowDateStr);

                    using (var frm = new AttachmentForm(currentVal, _dbName, _tableName, targetFolder, path => DeletePhysicalFile(path, e.RowIndex))) 
                    {
                        if (frm.ShowDialog() == DialogResult.OK) 
                        {
                            _dgv[e.ColumnIndex, e.RowIndex].Value = frm.FinalPathsString;
                            _dgv.EndEdit();
                        }
                    }
                }
            }
        }

        private bool IsFileUsedInDatabase(string relativePath)
        {
            try 
            {
                DataTable dt = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                foreach (DataRow row in dt.Rows) 
                {
                    string val = row["附件檔案"]?.ToString();
                    if (!string.IsNullOrEmpty(val) && val.Contains(relativePath)) return true;
                }
                return false;
            } 
            catch { return true; } 
        }

        private void DeletePhysicalFile(string relativePath, int currentRowIndex) 
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return;
            
            bool isUsedByOthers = false;
            foreach (DataGridViewRow row in _dgv.Rows) 
            {
                if (row.Index == currentRowIndex || row.IsNewRow) continue;
                if (_dgv.Columns.Contains("附件檔案")) 
                {
                    string cellVal = row.Cells["附件檔案"].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellVal)) 
                    {
                        string[] paths = cellVal.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (paths.Contains(relativePath)) { isUsedByOthers = true; break; }
                    }
                }
            }
            
            if (!isUsedByOthers && IsFileUsedInDatabase(relativePath)) isUsedByOthers = true;
            if (isUsedByOthers) return;
            
            try 
            {
                string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                if (File.Exists(absPath)) 
                {
                    File.Delete(absPath); 
                    DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(absPath));
                    string attachRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");
                    
                    while (dir != null && dir.FullName.StartsWith(attachRootDir) && dir.FullName.Length > attachRootDir.Length) 
                    {
                        if (dir.Exists && dir.GetFiles().Length == 0 && dir.GetDirectories().Length == 0) 
                        {
                            dir.Delete(); dir = dir.Parent;
                        } 
                        else break; 
                    }
                }
            } 
            catch { }
        }

        private void EnforceDateFormats(DataTable dt) 
        {
            if (dt == null || !dt.Columns.Contains(_dateColumnName)) return;
            
            string format = "yyyy-MM-dd";
            if (_timeMode == TimeMode.YearMonth) format = "yyyy-MM";
            else if (_timeMode == TimeMode.Year) format = "yyyy";
            
            foreach (DataRow row in dt.Rows) 
            {
                if (row.RowState == DataRowState.Deleted) continue;
                
                string val = row[_dateColumnName]?.ToString();
                
                if (!string.IsNullOrWhiteSpace(val)) 
                {
                    val = val.Replace("/", "-");
                    if (_timeMode == TimeMode.Year && val.Length == 4 && int.TryParse(val, out _)) 
                    {
                        row[_dateColumnName] = val; 
                        continue;
                    }
                    if (DateTime.TryParse(val, out DateTime d)) 
                    {
                        row[_dateColumnName] = d.ToString(format);
                    }
                }
            }
        }

        private string GetDateString(ComboBox y, ComboBox m, ComboBox d) 
        {
            if (_timeMode == TimeMode.Year) return y.SelectedItem.ToString();
            if (_timeMode == TimeMode.YearMonth) return $"{y.SelectedItem}-{m.SelectedItem}";
            
            return $"{y.SelectedItem}-{m.SelectedItem}-{d.SelectedItem}";
        }

        private void UpdateCboColumns() 
        {
            string currentSearchSel = _cboSearchColumn.SelectedItem?.ToString();
            
            _cboColumns.Items.Clear();
            _cboSearchColumn.Items.Clear();
            _cboSearchColumn.Items.Add(""); 

            foreach (DataGridViewColumn c in _dgv.Columns) 
            {
                if (c.Name != "Id" && c.Name != _dateColumnName) 
                {
                    _cboColumns.Items.Add(c.Name);
                }
                
                if (c.Name != "Id") 
                {
                    _cboSearchColumn.Items.Add(c.Name);
                }
            }

            if (!string.IsNullOrEmpty(currentSearchSel) && _cboSearchColumn.Items.Contains(currentSearchSel)) 
            {
                _cboSearchColumn.SelectedItem = currentSearchSel;
            } 
            else if (_cboSearchColumn.Items.Count > 0) 
            {
                _cboSearchColumn.SelectedIndex = 0;
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) 
        {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2"); 
            d.SelectedItem = date.Day.ToString("D2");
        }

        private void SyncAttachmentPaths(DataTable dt) 
        {
            if (!dt.Columns.Contains("附件檔案")) return;

            foreach (DataRow row in dt.Rows) 
            {
                if (row.RowState == DataRowState.Deleted) continue;
                string attachStr = row["附件檔案"]?.ToString();
                if (string.IsNullOrEmpty(attachStr)) continue;

                string rowDateStr = row[_dateColumnName]?.ToString() ?? "";
                string targetFolder = GetExpectedFolderName(rowDateStr); 

                string[] paths = attachStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                bool changed = false;
                
                for (int i = 0; i < paths.Length; i++) 
                {
                    string oldRelPath = paths[i].Replace("\\", "/");
                    string fileName = Path.GetFileName(oldRelPath);
                    string oldDir = Path.GetDirectoryName(oldRelPath).Replace("\\", "/");

                    string expectedRelDir = $"附件/{_dbName}/{_tableName}/{targetFolder}";

                    if (!oldDir.Equals(expectedRelDir, StringComparison.OrdinalIgnoreCase)) 
                    {
                        bool usedByOthersInGrid = false;
                        foreach(DataRow r in dt.Rows) {
                            if (r == row || r.RowState == DataRowState.Deleted) continue;
                            string otherAttach = r["附件檔案"]?.ToString();
                            if (!string.IsNullOrEmpty(otherAttach) && otherAttach.Contains(oldRelPath)) {
                                usedByOthersInGrid = true;
                                break;
                            }
                        }

                        bool usedByOthersInDb = false;
                        int currentRowId = -1;
                        if (dt.Columns.Contains("Id") && row["Id"] != DBNull.Value) {
                            int.TryParse(row["Id"].ToString(), out currentRowId);
                        }

                        try {
                            DataTable dbDt = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                            foreach (DataRow dbRow in dbDt.Rows) {
                                int dbId = Convert.ToInt32(dbRow["Id"]);
                                if (dbId == currentRowId) continue;
                                string dbAttach = dbRow["附件檔案"]?.ToString();
                                if (!string.IsNullOrEmpty(dbAttach) && dbAttach.Contains(oldRelPath)) {
                                    usedByOthersInDb = true;
                                    break;
                                }
                            }
                        } catch { }

                        if (usedByOthersInGrid || usedByOthersInDb) {
                            continue; 
                        }

                        string oldAbsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, oldRelPath);
                        string newAbsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, expectedRelDir);
                        if (!Directory.Exists(newAbsDir)) Directory.CreateDirectory(newAbsDir);

                        string newAbsPath = Path.Combine(newAbsDir, fileName);
                        
                        int counter = 1;
                        string baseName = Path.GetFileNameWithoutExtension(fileName);
                        string ext = Path.GetExtension(fileName);
                        while (File.Exists(newAbsPath) && oldAbsPath != newAbsPath) 
                        {
                            fileName = $"{baseName}_{counter++}{ext}";
                            newAbsPath = Path.Combine(newAbsDir, fileName);
                        }

                        if (File.Exists(oldAbsPath)) 
                        {
                            File.Move(oldAbsPath, newAbsPath);
                            paths[i] = $"{expectedRelDir}/{fileName}";
                            changed = true;
                        }
                    }
                }
                
                if (changed) 
                {
                    row["附件檔案"] = string.Join("|", paths);
                }
            }
        }

        private async void BtnSave_Click(object sender, EventArgs e) 
        {
            try 
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit(); 
                SaveColumnOrder(); 
                SetUIState(false, "資料庫寫入與檔案同步中，請稍候...", Color.Orange);
                
                DataTable dt = (DataTable)_dgv.DataSource;
                bool success = false;
                
                await Task.Run(() => { 
                    EnforceDateFormats(dt); 
                    SyncAttachmentPaths(dt);
                    success = DataManager.BulkSaveTable(_dbName, _tableName, dt); 
                });
                
                if (success) 
                { 
                    SetUIState(true, "資料儲存成功！", Color.Green); 
                    MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    
                    // 🟢 存檔後使用 ReloadCurrentDataAsync，保留使用者的查詢條件與 Scroll 位置
                    await ReloadCurrentDataAsync(); 
                } 
                else 
                { 
                    SetUIState(true, "資料儲存失敗", Color.Red); 
                }
            } 
            catch (Exception ex) 
            { 
                SetUIState(true, "儲存異常", Color.Red); 
                MessageBox.Show("儲存異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
            finally 
            { 
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
            }
        }

        private void BtnExport_Click(object sender, EventArgs e) 
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = _chineseTitle + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) 
                        { 
                            using (ExcelPackage p = new ExcelPackage()) 
                            { 
                                var ws = p.Workbook.Worksheets.Add("Data"); 
                                ws.Cells["A1"].LoadFromDataTable(dt, true); 
                                p.SaveAs(new FileInfo(sfd.FileName)); 
                            } 
                        } 
                        else 
                        {
                            StringBuilder sb = new StringBuilder(); 
                            sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            
                            foreach (DataRow r in dt.Rows) 
                            {
                                sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功！(附件欄位將輸出為相對路徑，以保證資料完整性)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
            }
        }

        private async void BtnImportExcel_Click(object sender, EventArgs e) 
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "請選擇要匯入的 Excel 檔案" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                        SetUIState(false, "Excel 解析與背景運算中，請稍候...", Color.Orange);

                        DataTable dt = (DataTable)_dgv.DataSource;
                        _dgv.DataSource = null; 
                        
                        await Task.Run(() => {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) 
                            {
                                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                if (ws == null || ws.Dimension == null) return;
                                
                                int rowCount = ws.Dimension.Rows; 
                                int colCount = ws.Dimension.Columns;
                                
                                string[] headers = new string[colCount];
                                for (int c = 1; c <= colCount; c++) 
                                {
                                    headers[c - 1] = ws.Cells[1, c].Text.Trim();
                                }

                                _calcHelper?.BeginBulkUpdate();
                                
                                for (int r = 2; r <= rowCount; r++) 
                                {
                                    DataRow nr = dt.NewRow(); 
                                    bool hasData = false;
                                    
                                    for (int c = 1; c <= colCount; c++) 
                                    {
                                        string cn = headers[c - 1]; 
                                        string val = ws.Cells[r, c].Text.Trim(); 
                                        
                                        if (dt.Columns.Contains(cn) && cn != "Id" && !string.IsNullOrEmpty(val)) 
                                        {
                                            nr[cn] = val; 
                                            hasData = true;
                                        }
                                    }
                                    if (hasData) dt.Rows.Add(nr);
                                }
                                
                                _calcHelper?.RecalculateTable(dt); 
                                _calcHelper?.EndBulkUpdate(); 
                                EnforceDateFormats(dt);
                            }
                        });
                        
                        _dgv.DataSource = dt; 
                        ApplyGridStyles(); 
                        RestoreColumnOrder();
                        SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{dt.Rows.Count}", Color.Green);
                        MessageBox.Show("Excel 匯入成功！\n請檢查數據後點擊「儲存數據」。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        await LoadGridDataAsync(); 
                        MessageBox.Show("匯入異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    } 
                    finally 
                    { 
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
                    }
                }
            }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e) 
        {
            if (e.Control && e.KeyCode == Keys.V) 
            {
                try 
                {
                    string text = Clipboard.GetText(); 
                    if (string.IsNullOrEmpty(text)) return;
                    
                    _calcHelper?.BeginBulkUpdate();
                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    
                    int r = _dgv.CurrentCell.RowIndex;
                    int c = _dgv.CurrentCell.ColumnIndex;
                    DataTable dt = (DataTable)_dgv.DataSource;
                    
                    foreach (string line in lines) 
                    {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++) 
                        {
                            if (c + i < _dgv.Columns.Count) 
                            {
                                if (_dgv.Columns[c + i].Name.Contains("附件檔案") || !_dgv.Columns[c + i].ReadOnly) 
                                {
                                    _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                                }
                            }
                        }
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt); 
                    _calcHelper?.EndBulkUpdate(); 
                    EnforceDateFormats(dt); 
                    _dgv.Refresh();
                } 
                catch 
                { 
                    _calcHelper?.EndBulkUpdate(); 
                }
            }
        }

        private class AttachmentForm : Form
        {
            public string FinalPathsString { get; private set; }
            private List<string> _paths = new List<string>();
            private string _dbName, _tableName, _targetFolder;
            private Action<string> _deleteAction;
            private FlowLayoutPanel _flpList;

            public AttachmentForm(string currentRelPathStr, string dbName, string tableName, string targetFolder, Action<string> deleteAction) 
            {
                _dbName = dbName; 
                _tableName = tableName; 
                _targetFolder = targetFolder; 
                _deleteAction = deleteAction;
                
                if (!string.IsNullOrEmpty(currentRelPathStr)) 
                {
                    _paths = new List<string>(currentRelPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries));
                }
                
                this.Text = "多檔附件管理中心"; 
                this.Size = new Size(700, 600); 
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.FixedDialog; 
                this.MaximizeBox = false; 
                this.MinimizeBox = false; 
                this.BackColor = Color.White;

                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); 
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));

                GroupBox boxList = new GroupBox { Text = "1. 已上傳檔案清單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
                _flpList = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                boxList.Controls.Add(_flpList); 
                tlp.Controls.Add(boxList, 0, 0);

                GroupBox boxUpload = new GroupBox { Text = "2. 新增附件檔案", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
                Panel pnlDrop = new Panel { Dock = DockStyle.Fill, AllowDrop = true, BackColor = Color.AliceBlue, Cursor = Cursors.Hand };
                pnlDrop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlDrop.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Dashed);
                
                Label lblDrop = new Label { Text = "📁 點擊此處選擇多個檔案\n\n或\n\n將檔案拖曳至此區域", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.SteelBlue };
                lblDrop.Click += (s, e) => SelectFiles(); 
                pnlDrop.Click += (s, e) => SelectFiles(); 
                pnlDrop.Controls.Add(lblDrop);
                
                pnlDrop.DragEnter += (s, e) => { 
                    if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; 
                };
                pnlDrop.DragDrop += (s, e) => { 
                    ProcessUpload((string[])e.Data.GetData(DataFormats.FileDrop)); 
                };
                boxUpload.Controls.Add(pnlDrop); 
                tlp.Controls.Add(boxUpload, 0, 1);

                Button btnClearAll = new Button { Text = "🗑️ 清除此筆紀錄的所有附件", Dock = DockStyle.Fill, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(3, 5, 3, 5) };
                btnClearAll.Click += (s, e) => {
                    if (_paths.Count == 0) return;
                    if (MessageBox.Show("確定要清除所有附件嗎？\n(實體檔案將被同步永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                    {
                        foreach (var p in _paths) _deleteAction(p);
                        _paths.Clear(); 
                        RefreshListUI();
                    }
                };
                tlp.Controls.Add(btnClearAll, 0, 2);

                Button btnSaveClose = new Button { Text = "💾 確認變更並返回", Dock = DockStyle.Fill, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Margin = new Padding(3, 5, 3, 5) };
                btnSaveClose.Click += (s, e) => { 
                    FinalPathsString = string.Join("|", _paths); 
                    this.DialogResult = DialogResult.OK; 
                };
                tlp.Controls.Add(btnSaveClose, 0, 3);

                this.Controls.Add(tlp); 
                RefreshListUI();
            }

            private void RefreshListUI() 
            {
                _flpList.Controls.Clear();
                if (_paths.Count == 0) 
                { 
                    _flpList.Controls.Add(new Label { Text = "(尚無任何附件)", ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(10) }); 
                    return; 
                }
                
                foreach (string path in _paths) 
                {
                    Panel pItem = new Panel { Width = _flpList.Width - 30, Height = 40, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                    Label lName = new Label { Text = Path.GetFileName(path), Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Microsoft JhengHei UI", 11F) };
                    
                    Button bOpen = new Button { Text = "開啟", Width = 100, Dock = DockStyle.Right, BackColor = Color.LightGray, Cursor = Cursors.Hand };
                    bOpen.Click += (s, e) => { 
                        try { System.Diagnostics.Process.Start(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path)); } 
                        catch (Exception ex) { MessageBox.Show("開啟失敗：" + ex.Message); } 
                    };

                    Button bDownload = new Button { Text = "下載", Width = 100, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
                    bDownload.Click += (s, e) => {
                        try 
                        {
                            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
                            if (!File.Exists(sourcePath)) 
                            {
                                MessageBox.Show("找不到原始檔案，可能已被移動或刪除。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            using (SaveFileDialog sfd = new SaveFileDialog())
                            {
                                string fileName = Path.GetFileName(path);
                                string ext = Path.GetExtension(path);
                                sfd.FileName = fileName;
                                sfd.Title = "另存附件檔案";
                                sfd.Filter = $"檔案 (*{ext})|*{ext}|所有檔案 (*.*)|*.*";
                                
                                if (sfd.ShowDialog() == DialogResult.OK)
                                {
                                    File.Copy(sourcePath, sfd.FileName, true);
                                    MessageBox.Show("檔案下載完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        catch (Exception ex) 
                        {
                            MessageBox.Show("下載失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };
                    
                    Button bDel = new Button { Text = "刪除", Width = 100, Dock = DockStyle.Right, BackColor = Color.LightCoral, ForeColor = Color.White, Cursor = Cursors.Hand };
                    bDel.Click += (s, e) => { 
                        if (MessageBox.Show($"確定刪除 {Path.GetFileName(path)}?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                        { 
                            _deleteAction(path); 
                            _paths.Remove(path); 
                            RefreshListUI(); 
                        } 
                    };
                    
                    pItem.Controls.Add(lName); 
                    pItem.Controls.Add(bDel);       
                    pItem.Controls.Add(bDownload);  
                    pItem.Controls.Add(bOpen);      
                    
                    _flpList.Controls.Add(pItem);
                }
            }

            private void SelectFiles() 
            {
                using (OpenFileDialog ofd = new OpenFileDialog { Title = "選擇附件檔案", Multiselect = true, Filter = "所有檔案 (*.*)|*.*" }) 
                {
                    if (ofd.ShowDialog() == DialogResult.OK) ProcessUpload(ofd.FileNames);
                }
            }

            private void ProcessUpload(string[] sourceFiles) 
            {
                if (sourceFiles.Length == 0) return;
                
                using (ImageCompressionHelper compressor = new ImageCompressionHelper())
                {
                    string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件", _dbName, _tableName, _targetFolder);
                    
                    if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
                    
                    foreach (string src in sourceFiles) 
                    {
                        try 
                        {
                            string ext = Path.GetExtension(src); 
                            string baseName = Path.GetFileNameWithoutExtension(src);
                            string destName = baseName + ext; 
                            string destPath = Path.Combine(destDir, destName);
                            
                            int count = 1; 
                            while (File.Exists(destPath)) 
                            { 
                                destName = $"{baseName}_{count++}{ext}"; 
                                destPath = Path.Combine(destDir, destName); 
                            }
                            
                            compressor.ProcessAndSave(src, destPath);
                            
                            _paths.Add($"附件/{_dbName}/{_tableName}/{_targetFolder}/{destName}");
                        } 
                        catch (Exception ex) 
                        { 
                            MessageBox.Show($"上傳檔案 {Path.GetFileName(src)} 失敗: {ex.Message}", "錯誤"); 
                        }
                    }
                }
                RefreshListUI();
            }
        }
        
        private class ImageCompressionHelper : IDisposable
        {
            private readonly string[] _imageExts = { ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
            private ImageCodecInfo _jpgEncoder;
            private EncoderParameters _encoderParams;

            public ImageCompressionHelper()
            {
                _jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                _encoderParams = new EncoderParameters(1);
                _encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L); // 100% 最高畫質
            }

            public void ProcessAndSave(string srcPath, string destPath)
            {
                string ext = Path.GetExtension(srcPath).ToLower();

                if (!_imageExts.Contains(ext))
                {
                    File.Copy(srcPath, destPath);
                    return;
                }

                using (Image originalImg = Image.FromFile(srcPath))
                {
                    int maxSide = 1024;
                    int origWidth = originalImg.Width;
                    int origHeight = originalImg.Height;

                    if (origWidth > maxSide || origHeight > maxSide)
                    {
                        float ratio = Math.Min((float)maxSide / origWidth, (float)maxSide / origHeight);
                        int newWidth = (int)(origWidth * ratio);
                        int newHeight = (int)(origHeight * ratio);

                        using (Bitmap resizedImg = new Bitmap(newWidth, newHeight))
                        {
                            using (Graphics g = Graphics.FromImage(resizedImg))
                            {
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = SmoothingMode.HighQuality;
                                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                g.CompositingQuality = CompositingQuality.HighQuality;
                                
                                g.DrawImage(originalImg, 0, 0, newWidth, newHeight);
                            }

                            if ((ext == ".jpg" || ext == ".jpeg") && _jpgEncoder != null)
                            {
                                resizedImg.Save(destPath, _jpgEncoder, _encoderParams);
                            }
                            else
                            {
                                resizedImg.Save(destPath, originalImg.RawFormat);
                            }
                        }
                    }
                    else
                    {
                        File.Copy(srcPath, destPath);
                    }
                }
            }

            private ImageCodecInfo GetEncoder(ImageFormat format)
            {
                ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
                return codecs.FirstOrDefault(codec => codec.FormatID == format.Guid);
            }

            public void Dispose()
            {
                _encoderParams?.Dispose();
            }
        }
    }
}
