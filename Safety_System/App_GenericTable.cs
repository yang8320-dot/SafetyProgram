/// FILE: Safety_System/App_GenericTable.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
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

        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        
        private Button _btnToggle, _btnRead, _btnSave, _btnExport, _btnImport;
        private Label _lblStatus;     

        private ComboBox _cboSearchColumn;
        private TextBox _txtSearchKeyword;
        private Button _btnAdvancedSearch;

        private bool _isFirstLoad = true;
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

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
            // 🟢 呼叫 TableSchemaManager 來獲取資料表結構
            string schema = TableSchemaManager.SchemaMap.ContainsKey(_tableName) ? TableSchemaManager.SchemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

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

            if (columns.Contains("日期")) { _timeMode = TimeMode.Date; _dateColumnName = "日期"; }
            else if (columns.Contains("年月")) { _timeMode = TimeMode.YearMonth; _dateColumnName = "年月"; }
            else if (columns.Contains("年度")) { _timeMode = TimeMode.Year; _dateColumnName = "年度"; }
            else 
            {
                string altDate = columns.FirstOrDefault(c => c.Contains("日期"));
                if (altDate != null) { _timeMode = TimeMode.Date; _dateColumnName = altDate; }
                else 
                {
                    altDate = columns.FirstOrDefault(c => c.Contains("年月"));
                    if (altDate != null) { _timeMode = TimeMode.YearMonth; _dateColumnName = altDate; }
                    else 
                    {
                        altDate = columns.FirstOrDefault(c => c.Contains("年度"));
                        if (altDate != null) { _timeMode = TimeMode.Year; _dateColumnName = altDate; }
                        else { _timeMode = TimeMode.Date; _dateColumnName = columns.FirstOrDefault(c => c != "Id") ?? "Id"; }
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

            if (_timeMode == TimeMode.YearMonth || _timeMode == TimeMode.Year) {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-6));
            } else {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            }
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            _btnRead = new Button { Text = "🔍 讀取資料", Size = new Size(150, 35), BackColor = Color.WhiteSmoke };
            _btnRead.Click += async (s, e) => { _isFirstLoad = false; await LoadGridDataAsync(); };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(150, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click; 
            
            _btnExport = new Button { Text = "📤 匯出Excel", Size = new Size(150, 35) }; 
            _btnExport.Click += BtnExport_Click;

            _btnImport = new Button { Text = "📥 匯入Excel", Size = new Size(150, 35) }; 
            _btnImport.Click += BtnImportExcel_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
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

            if (_timeMode == TimeMode.YearMonth) {
                _cboStartDay.Visible = false; lblSD.Visible = false; _cboEndDay.Visible = false; lblED.Visible = false;
            } else if (_timeMode == TimeMode.Year) {
                _cboStartDay.Visible = false; lblSD.Visible = false; _cboEndDay.Visible = false; lblED.Visible = false;
                _cboStartMonth.Visible = false; lblSM.Visible = false; _cboEndMonth.Visible = false; lblEM.Visible = false;
            }

            row1.Controls.AddRange(new Control[] {
                lblRange, _cboStartYear, lblSY, _cboStartMonth, lblSM, _cboStartDay, lblSD,
                lblTilde, _cboEndYear, lblEY, _cboEndMonth, lblEM, _cboEndDay, lblED,
                _btnRead, _btnExport, _btnImport, _btnToggle, _btnSave
            });

            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 10) };
            FlowLayoutPanel flpAdv = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            FlowLayoutPanel rowAdv1 = new FlowLayoutPanel { AutoSize = true };
            _txtNewColName = new TextBox { Width = 150 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(120, 35) };
            bAdd.Click += async (s, e) => { 
                if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) 
                { DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); await LoadGridDataAsync(); _txtNewColName.Clear(); } 
            };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; 
            _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "修改名稱", Size = new Size(120, 35) };
            bRen.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) 
                { DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); await LoadGridDataAsync(); _txtRenameCol.Clear(); } 
            };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { 
                    if(MessageBox.Show($"確定刪除整欄【{_cboColumns.SelectedItem}】？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    { DataManager.DropColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString()); await LoadGridDataAsync(); } 
                } 
            };
            
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(140, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += async (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>()
                                       .Select(c => c.OwningRow)
                                       .Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value)
                                       .Distinct().ToList();
                                       
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(包含所屬的實體附件檔案也將被永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    if (AuthManager.VerifyUser()) {
                        foreach (var r in selectedRows) {
                            if (_dgv.Columns.Contains("附件檔案")) {
                                string relPathStr = r.Cells["附件檔案"].Value?.ToString();
                                if (!string.IsNullOrEmpty(relPathStr)) {
                                    string[] paths = relPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var p in paths) DeletePhysicalFile(p, r.Index);
                                }
                            }
                            DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                        }
                        await LoadGridDataAsync(); 
                        MessageBox.Show("刪除成功！");
                    }
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0), WrapContents = false };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(140, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bLimitRead.Click += async (s, e) => { 
                if (int.TryParse(txtLimit.Text, out int l)) { 
                    SetUIState(false, "讀取中...", Color.Orange);
                    DataTable dt = null;
                    await Task.Run(() => { dt = DataManager.GetLatestRecords(_dbName, _tableName, l); EnforceDateFormats(dt); });
                    _dgv.DataSource = dt; 
                    ApplyGridStyles(); 
                    RestoreColumnOrder();
                    SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
                } 
            };

            _cboSearchColumn = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtSearchKeyword = new TextBox { Width = 180 };
            _btnAdvancedSearch = new Button { Text = "🔍 條件搜尋", Size = new Size(140, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            _btnAdvancedSearch.Click += async (s, e) => await ExecuteAdvancedSearchAsync();
            
            rowAdv2.Controls.AddRange(new Control[] { 
                new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead,
                new Label { Text = "查詢資料:", AutoSize = true, Margin = new Padding(40, 8, 0, 0) }, _cboSearchColumn, 
                new Label { Text = "關鍵字(包含):", AutoSize = true, Margin = new Padding(15, 8, 0, 0) }, _txtSearchKeyword, _btnAdvancedSearch 
            });

            flpAdv.Controls.Add(rowAdv1); 
            flpAdv.Controls.Add(rowAdv2);
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
            
            // 🟢 防止資料行在下拉選單找不到對應數值時報錯
            _dgv.DataError += (s, e) => { e.ThrowException = false; };
            
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
            _btnAdvancedSearch.Enabled = isEnabled;
            
            _lblStatus.Text = statusText; 
            _lblStatus.ForeColor = statusColor;
        }

        private void ApplyGridStyles() 
        {
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
                if (col.Name.Contains("附件檔案")) 
                {
                    col.ReadOnly = true; 
                    col.DefaultCellStyle.ForeColor = Color.Blue;
                    col.DefaultCellStyle.Font = new Font(_dgv.Font, FontStyle.Underline);
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }

            // 🟢 呼叫設定下拉式選單的核心邏輯
            SetupDropdownColumns();
        }

        // 🟢 將一般欄位自動替換為下拉式選單 (如果 TableSchemaManager 有定義)
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

                    List<string> finalItems = new List<string>(items);
                    
                    // 防呆設計：把已經存在於資料庫但「不在預設清單內」的舊有資料加入，避免 DataError 報錯
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
                        if (parts.Length > 1) e.Value = $"📁 [共 {parts.Length} 個檔案]";
                        else e.Value = Path.GetFileName(parts[0]);
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
                    if (_timeMode == TimeMode.Year && val.Length == 4 && int.TryParse(val, out _)) { row[_dateColumnName] = val; continue; }
                    if (DateTime.TryParse(val, out DateTime d)) row[_dateColumnName] = d.ToString(format);
                }
            }
        }

        private async Task LoadGridDataAsync() 
        {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
            string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);

            await Task.Run(() => {
                if (_isFirstLoad) { dt = DataManager.GetLatestRecords(_dbName, _tableName, 30); _isFirstLoad = false; } 
                else { dt = DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate); }
                EnforceDateFormats(dt);
            });

            _dgv.DataSource = dt;
            ApplyGridStyles();
            UpdateCboColumns();
            RestoreColumnOrder();

            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
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
                if (!string.IsNullOrEmpty(searchCol) && !string.IsNullOrWhiteSpace(keyword)) 
                {
                    dv.RowFilter = $"[{searchCol}] LIKE '%{keyword.Replace("'", "''")}%'";
                }
                dv.Sort = "Id DESC"; 
                resultDt = dv.ToTable(); 
                EnforceDateFormats(resultDt);
            });

            _dgv.DataSource = resultDt;
            ApplyGridStyles();
            RestoreColumnOrder();
            SetUIState(true, $"搜尋完成，共找到 {resultDt.Rows.Count} 筆符合條件資料", Color.Green);
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
            _cboColumns.Items.Clear(); _cboSearchColumn.Items.Clear(); _cboSearchColumn.Items.Add(""); 

            foreach (DataGridViewColumn c in _dgv.Columns) 
            {
                if (c.Name != "Id" && c.Name != _dateColumnName) _cboColumns.Items.Add(c.Name);
                if (c.Name != "Id") _cboSearchColumn.Items.Add(c.Name);
            }

            if (!string.IsNullOrEmpty(currentSearchSel) && _cboSearchColumn.Items.Contains(currentSearchSel)) _cboSearchColumn.SelectedItem = currentSearchSel;
            else if (_cboSearchColumn.Items.Count > 0) _cboSearchColumn.SelectedIndex = 0;
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) 
        {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2"); d.SelectedItem = date.Day.ToString("D2");
        }

        private void SaveColumnOrder() 
        { 
            try { var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); File.WriteAllText($"ColOrder_{_dbName}_{_tableName}.txt", string.Join(",", ordered), Encoding.UTF8); } 
            catch { } 
        }
        
        private void RestoreColumnOrder() 
        { 
            try { string fn = $"ColOrder_{_dbName}_{_tableName}.txt"; if (File.Exists(fn)) { string[] saved = File.ReadAllText(fn, Encoding.UTF8).Split(','); for (int i = 0; i < saved.Length; i++) { if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; } } } 
            catch { } 
        }

        private void SyncAttachmentPaths(DataTable dt) 
        {
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
                            if (!string.IsNullOrEmpty(otherAttach) && otherAttach.Contains(oldRelPath)) { usedByOthersInGrid = true; break; }
                        }

                        bool usedByOthersInDb = false;
                        int currentRowId = -1;
                        if (dt.Columns.Contains("Id") && row["Id"] != DBNull.Value) int.TryParse(row["Id"].ToString(), out currentRowId);

                        try {
                            DataTable dbDt = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                            foreach (DataRow dbRow in dbDt.Rows) {
                                int dbId = Convert.ToInt32(dbRow["Id"]);
                                if (dbId == currentRowId) continue;
                                string dbAttach = dbRow["附件檔案"]?.ToString();
                                if (!string.IsNullOrEmpty(dbAttach) && dbAttach.Contains(oldRelPath)) { usedByOthersInDb = true; break; }
                            }
                        } catch { }

                        if (usedByOthersInGrid || usedByOthersInDb) continue; 

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
                if (changed) row["附件檔案"] = string.Join("|", paths);
            }
        }

        private async void BtnSave_Click(object sender, EventArgs e) 
        {
            try {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit(); SaveColumnOrder(); SetUIState(false, "資料庫寫入與檔案同步中，請稍候...", Color.Orange);
                
                DataTable dt = (DataTable)_dgv.DataSource;
                bool success = false;
                await Task.Run(() => { EnforceDateFormats(dt); SyncAttachmentPaths(dt); success = DataManager.BulkSaveTable(_dbName, _tableName, dt); });
                
                if (success) { SetUIState(true, "資料儲存成功！", Color.Green); MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); await LoadGridDataAsync(); } 
                else { SetUIState(true, "資料儲存失敗", Color.Red); }
            } 
            catch (Exception ex) { SetUIState(true, "儲存異常", Color.Red); MessageBox.Show("儲存異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; }
        }

        private void BtnExport_Click(object sender, EventArgs e) 
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = _chineseTitle + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) { 
                            using (ExcelPackage p = new ExcelPackage()) { var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable(dt, true); p.SaveAs(new FileInfo(sfd.FileName)); } 
                        } else {
                            StringBuilder sb = new StringBuilder(); sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            foreach (DataRow r in dt.Rows) sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功！(附件欄位將輸出為相對路徑，以保證資料完整性)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
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
                                
                                int rowCount = ws.Dimension.Rows; int colCount = ws.Dimension.Columns;
                                string[] headers = new string[colCount];
                                for (int c = 1; c <= colCount; c++) headers[c - 1] = ws.Cells[1, c].Text.Trim();

                                _calcHelper?.BeginBulkUpdate();
                                for (int r = 2; r <= rowCount; r++) {
                                    DataRow nr = dt.NewRow(); bool hasData = false;
                                    for (int c = 1; c <= colCount; c++) {
                                        string cn = headers[c - 1]; string val = ws.Cells[r, c].Text.Trim(); 
                                        if (dt.Columns.Contains(cn) && cn != "Id" && !string.IsNullOrEmpty(val)) { nr[cn] = val; hasData = true; }
                                    }
                                    if (hasData) dt.Rows.Add(nr);
                                }
                                _calcHelper?.RecalculateTable(dt); _calcHelper?.EndBulkUpdate(); EnforceDateFormats(dt);
                            }
                        });
                        
                        _dgv.DataSource = dt; ApplyGridStyles(); RestoreColumnOrder();
                        SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{dt.Rows.Count}", Color.Green);
                        MessageBox.Show("Excel 匯入成功！\n請檢查數據後點擊「儲存數據」。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { await LoadGridDataAsync(); MessageBox.Show("匯入異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); } 
                    finally { if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; }
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
                            if (c + i < _dgv.Columns.Count && (_dgv.Columns[c + i].Name.Contains("附件檔案") || !_dgv.Columns[c + i].ReadOnly)) {
                                _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                            }
                        }
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt); _calcHelper?.EndBulkUpdate(); EnforceDateFormats(dt); _dgv.Refresh();
                } catch { _calcHelper?.EndBulkUpdate(); }
            }
        }

        // 附件管理 Form 及圖片壓縮 Helper 省略未更動，直接保留原始結構
        // ... (保持原樣，因為長度限制，這裡我只放類別開頭，但您原始程式碼可以直接保留)
        private class AttachmentForm : Form { /* ... 與原先相同 ... */ }
        private class ImageCompressionHelper : IDisposable { /* ... 與原先相同 ... */ }
    }
}
