using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public partial class App_CoreTable
    {
        // ================= 全域變數區 =================
        private enum TimeMode { Date, YearMonth, Year }
        private TimeMode _timeMode = TimeMode.Date;

        private enum SearchMode { DateRange, Limit, Advanced }
        private SearchMode _currentSearchMode = SearchMode.DateRange;
        private int _currentLimit = 50; 

        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        
        private Button _btnToggle, _btnRead, _btnSave, _btnExport, _btnImport, _btnExportPdf, _btnColSettings;
        private Button _btnRtfToExcel; 
        private Label _lblStatus;     

        private ComboBox _cboSearchColumn;
        private TextBox _txtSearchKeyword;
        private TextBox _txtLatestCount; 
        private Button _btnAdvancedSearch;

        private bool _isFirstLoad = true;
        private bool _isApplyingWidths = false; 
        private bool _isCascading = false;
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        private readonly ITableLogic _logic; 
        
        private string _dateColumnName = "日期";
        private DataGridViewAutoCalcHelper _calcHelper; 

        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();
        private Dictionary<string, int> _columnWidths = new Dictionary<string, int>();

        private ContextMenuStrip _ctxMenu;
        private int _rightClickedColIndex = -1;
        private string _frozenColumnName = null;

        // ================= 建構子與進入點 =================
        public App_CoreTable(string dbName, string tableName, string chineseTitle, ITableLogic logic)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
            _logic = logic ?? new DefaultLogic(); 
        }

        public Control GetView()
        {
            // 1. 初始化資料表結構
            string schema = TableSchemaManager.SchemaMap.ContainsKey(_tableName) 
                            ? TableSchemaManager.SchemaMap[_tableName] 
                            : TableSchemaManager.DefaultCustomSchema;

            DataManager.InitTable(_dbName, _tableName, $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});");
            _logic.InitializeSchema(_dbName, _tableName);

            // 2. 載入記憶設定
            LoadVisibilitySettings();
            LoadColumnWidths(); 
            CheckTimeMode();

            // 3. 呼叫 UI 構建與事件綁定 (實作於 App_CoreTable.UI.cs 與 App_CoreTable.Events.cs)
            Control mainPanel = BuildUI();
            BindEvents();

            // 4. 載入初始資料
            _ = ReloadCurrentDataAsync(); 
            return mainPanel;
        }

        // ================= 核心資料邏輯 =================
        private void CheckTimeMode()
        {
            List<string> columns = DataManager.GetColumnNames(_dbName, _tableName);
            if (columns.Contains("月份")) {
                try { DataManager.RenameColumn(_dbName, _tableName, "月份", "年月"); columns = DataManager.GetColumnNames(_dbName, _tableName); } catch { }
            }

            if (columns.Contains("日期")) { _timeMode = TimeMode.Date; _dateColumnName = "日期"; }
            else if (columns.Contains("年月")) { _timeMode = TimeMode.YearMonth; _dateColumnName = "年月"; }
            else if (columns.Contains("年度")) { _timeMode = TimeMode.Year; _dateColumnName = "年度"; }
            else {
                _dateColumnName = columns.FirstOrDefault(c => c.Contains("日期")) ?? 
                                  columns.FirstOrDefault(c => c.Contains("年月")) ?? 
                                  columns.FirstOrDefault(c => c.Contains("年度")) ?? "Id";
                if (_dateColumnName.Contains("年月")) _timeMode = TimeMode.YearMonth;
                else if (_dateColumnName.Contains("年度")) _timeMode = TimeMode.Year;
                else _timeMode = TimeMode.Date;
            }
        }

        private string GetExpectedFolderName(string rowDateStr)
        {
            if (string.IsNullOrWhiteSpace(rowDateStr)) return DateTime.Now.ToString("yyyy-MM");
            if (_timeMode == TimeMode.Year && rowDateStr.Length >= 4) return rowDateStr.Substring(0, 4);
            if (rowDateStr.Length >= 7) return rowDateStr.Substring(0, 7);
            return DateTime.Now.ToString("yyyy-MM");
        }

        private async Task ReloadCurrentDataAsync()
        {
            int firstRowIndex = -1, selectedRowIndex = -1, selectedColIndex = -1;
            
            if (_dgv.Rows.Count > 0 && !_isFirstLoad) {
                firstRowIndex = _dgv.FirstDisplayedScrollingRowIndex;
                if (_dgv.CurrentCell != null) { 
                    selectedRowIndex = _dgv.CurrentCell.RowIndex; 
                    selectedColIndex = _dgv.CurrentCell.ColumnIndex; 
                }
            }

            if (_currentSearchMode == SearchMode.Advanced) { await ExecuteAdvancedSearchAsync(); }
            else if (_currentSearchMode == SearchMode.Limit) { await LoadLimitDataAsync(_currentLimit); }
            else { await LoadGridDataAsync(); }

            try {
                if (firstRowIndex >= 0 && firstRowIndex < _dgv.Rows.Count) _dgv.FirstDisplayedScrollingRowIndex = firstRowIndex;
                if (selectedRowIndex >= 0 && selectedRowIndex < _dgv.Rows.Count && selectedColIndex >= 0) {
                    _dgv.ClearSelection();
                    _dgv.CurrentCell = _dgv.Rows[selectedRowIndex].Cells[selectedColIndex];
                    _dgv.Rows[selectedRowIndex].Selected = true;
                }
            } catch { } 
        }

        private async Task LoadGridDataAsync() 
        {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
            string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);

            await Task.Run(() => {
                dt = _isFirstLoad ? DataManager.GetLatestRecords(_dbName, _tableName, 30) : DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate);
                EnforceDateFormats(dt);
            });

            UnfreezeAllColumns(); 
            _isApplyingWidths = true;
            _dgv.SuspendLayout();
            
            PreFillComboBoxItems(dt); 
            _dgv.DataSource = dt;
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();
            ApplyFreezeState(); 
            
            _dgv.ResumeLayout(true);
            _isApplyingWidths = false;
            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
        }

        private async Task LoadLimitDataAsync(int limit) 
        {
            SetUIState(false, "讀取中...", Color.Orange);
            DataTable dt = null;
            await Task.Run(() => { dt = DataManager.GetLatestRecords(_dbName, _tableName, limit); EnforceDateFormats(dt); });
            
            UnfreezeAllColumns();
            _isApplyingWidths = true;
            _dgv.SuspendLayout();
            
            PreFillComboBoxItems(dt); 
            _dgv.DataSource = dt; 
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();
            ApplyFreezeState();

            _dgv.ResumeLayout(true);
            _isApplyingWidths = false;
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

                if (!string.IsNullOrEmpty(searchCol)) {
                    if (keyword == "有鍵入資料者") dv.RowFilter = $"[{searchCol}] <> '' AND [{searchCol}] IS NOT NULL";
                    else if (string.IsNullOrWhiteSpace(keyword)) dv.RowFilter = $"[{searchCol}] IS NULL OR [{searchCol}] = ''";
                    else dv.RowFilter = $"[{searchCol}] LIKE '%{keyword.Replace("'", "''")}%'";
                }
                
                dv.Sort = "Id DESC"; 
                resultDt = dv.ToTable(); 
                
                if (_logic is LawLogic && int.TryParse(_txtLatestCount.Text, out int limit)) {
                    DataTable limitedDt = resultDt.Clone();
                    for (int i = 0; i < Math.Min(limit, resultDt.Rows.Count); i++) limitedDt.ImportRow(resultDt.Rows[i]);
                    resultDt = limitedDt;
                }
                EnforceDateFormats(resultDt);
            });

            UnfreezeAllColumns();
            _isApplyingWidths = true;
            _dgv.SuspendLayout();

            PreFillComboBoxItems(resultDt); 
            _dgv.DataSource = resultDt;
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();
            ApplyFreezeState();

            _dgv.ResumeLayout(true);
            _isApplyingWidths = false;
            SetUIState(true, $"搜尋完成，共找到 {resultDt.Rows.Count} 筆資料", Color.Green);
        }

        // ================= 檔案與記憶功能 =================
        private void LoadColumnWidths() {
            _columnWidths.Clear();
            var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Width");
            foreach (var kvp in dict) if (int.TryParse(kvp.Value, out int w)) _columnWidths[kvp.Key] = w; 
        }

        private void SaveColumnWidths() {
            foreach (var kvp in _columnWidths) DataManager.SaveGridConfig(_dbName, _tableName, "Width", kvp.Key, kvp.Value.ToString());
        }

        private void LoadVisibilitySettings() {
            _columnVisibility.Clear();
            var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Visibility");
            foreach (var kvp in dict) _columnVisibility[kvp.Key] = (kvp.Value == "1");
        }

        private void SaveVisibilitySettings() {
            DataManager.ClearGridConfig(_dbName, _tableName, "Visibility");
            foreach (var kvp in _columnVisibility) DataManager.SaveGridConfig(_dbName, _tableName, "Visibility", kvp.Key, kvp.Value ? "1" : "0");
        }

        private void SaveColumnOrder() { 
            try { 
                var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); 
                DataManager.SaveGridConfig(_dbName, _tableName, "Order", "All", string.Join(",", ordered)); 
            } catch { } 
        }
        
        private void RestoreColumnOrder() { 
            try { 
                UnfreezeAllColumns(); 
                var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Order"); 
                if (dict.ContainsKey("All")) { 
                    string[] saved = dict["All"].Split(','); 
                    for (int i = 0; i < saved.Length; i++) { 
                        if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; 
                    } 
                } 
            } catch { } 
        }

        private void SyncAttachmentPaths(DataTable dt) 
        {
            if (!dt.Columns.Contains("附件檔案")) return;
            
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                string attachStr = row["附件檔案"]?.ToString();
                if (string.IsNullOrEmpty(attachStr)) continue;
                
                string rowDateStr = row[_dateColumnName]?.ToString() ?? "";
                string targetFolder = GetExpectedFolderName(rowDateStr); 
                string[] paths = attachStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                bool changed = false;
                
                for (int i = 0; i < paths.Length; i++) {
                    string oldRelPath = paths[i].Replace("\\", "/"); 
                    string fileName = Path.GetFileName(oldRelPath); 
                    string oldDir = Path.GetDirectoryName(oldRelPath).Replace("\\", "/");
                    string expectedRelDir = $"附件/{_dbName}/{_tableName}/{targetFolder}";
                    
                    if (!oldDir.Equals(expectedRelDir, StringComparison.OrdinalIgnoreCase)) {
                        bool usedByOthersInGrid = false;
                        foreach(DataRow r in dt.Rows) { 
                            if (r == row || r.RowState == DataRowState.Deleted) continue; 
                            string otherAttach = r["附件檔案"]?.ToString(); 
                            if (!string.IsNullOrEmpty(otherAttach) && otherAttach.Contains(oldRelPath)) { usedByOthersInGrid = true; break; } 
                        }
                        
                        if (usedByOthersInGrid) continue; 
                        
                        string oldAbsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, oldRelPath);
                        string newAbsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, expectedRelDir);
                        if (!Directory.Exists(newAbsDir)) Directory.CreateDirectory(newAbsDir);
                        
                        string newAbsPath = Path.Combine(newAbsDir, fileName);
                        int counter = 1; 
                        string baseName = Path.GetFileNameWithoutExtension(fileName); 
                        string ext = Path.GetExtension(fileName);
                        
                        while (File.Exists(newAbsPath) && oldAbsPath != newAbsPath) { 
                            fileName = $"{baseName}_{counter++}{ext}"; 
                            newAbsPath = Path.Combine(newAbsDir, fileName); 
                        }
                        
                        if (File.Exists(oldAbsPath)) { 
                            File.Move(oldAbsPath, newAbsPath); 
                            paths[i] = $"{expectedRelDir}/{fileName}"; 
                            changed = true; 
                        }
                    }
                }
                if (changed) row["附件檔案"] = string.Join("|", paths);
            }
        }

        private void DeletePhysicalFile(string relativePath, int currentRowIndex) 
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return;
            bool isUsedByOthers = false;
            
            foreach (DataGridViewRow row in _dgv.Rows) {
                if (row.Index == currentRowIndex || row.IsNewRow) continue;
                if (_dgv.Columns.Contains("附件檔案")) {
                    string cellVal = row.Cells["附件檔案"].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellVal)) { 
                        string[] paths = cellVal.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries); 
                        if (paths.Contains(relativePath)) { isUsedByOthers = true; break; } 
                    }
                }
            }
            
            if (!isUsedByOthers) {
                try { 
                    DataTable dt = DataManager.GetTableData(_dbName, _tableName, "", "", ""); 
                    foreach (DataRow row in dt.Rows) { 
                        string val = row["附件檔案"]?.ToString(); 
                        if (!string.IsNullOrEmpty(val) && val.Contains(relativePath)) { isUsedByOthers = true; break; } 
                    } 
                } catch { } 
            }
            
            if (isUsedByOthers) return;
            
            try {
                string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                if (File.Exists(absPath)) {
                    File.Delete(absPath); 
                    DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(absPath));
                    string attachRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");
                    
                    while (dir != null && dir.FullName.StartsWith(attachRootDir) && dir.FullName.Length > attachRootDir.Length) { 
                        if (dir.Exists && dir.GetFiles().Length == 0 && dir.GetDirectories().Length == 0) { 
                            dir.Delete(); dir = dir.Parent; 
                        } else break; 
                    }
                }
            } catch { }
        }

        private void EnforceDateFormats(DataTable dt) 
        {
            if (dt == null || !dt.Columns.Contains(_dateColumnName)) return;
            string format = "yyyy-MM-dd";
            if (_timeMode == TimeMode.YearMonth) format = "yyyy-MM"; 
            else if (_timeMode == TimeMode.Year) format = "yyyy";
            
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                string val = row[_dateColumnName]?.ToString();
                if (!string.IsNullOrWhiteSpace(val)) {
                    val = val.Replace("/", "-");
                    if (_timeMode == TimeMode.Year && val.Length == 4 && int.TryParse(val, out _)) { row[_dateColumnName] = val; continue; }
                    if (DateTime.TryParse(val, out DateTime d)) { row[_dateColumnName] = d.ToString(format); }
                }
            }
        }

        private string GetDateString(ComboBox y, ComboBox m, ComboBox d) 
        {
            if (_timeMode == TimeMode.Year) return y.SelectedItem.ToString();
            if (_timeMode == TimeMode.YearMonth) return $"{y.SelectedItem}-{m.SelectedItem}";
            return $"{y.SelectedItem}-{m.SelectedItem}-{d.SelectedItem}";
        }
    }
}
