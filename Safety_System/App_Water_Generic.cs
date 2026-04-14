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
        // 🟢 定義時間模式：日期 (年月日)、年月 (隱藏日)
        private enum TimeMode { Date, YearMonth }
        private TimeMode _timeMode = TimeMode.Date;

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
        
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

        // 🟢 水資源五大表預設結構 (月報表已修正為[年月])
        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            { "WaterMeterReadings", "[日期] TEXT, [星期] TEXT, [用電量] TEXT, [用電量日統計] TEXT, [廢水進流量] TEXT, [廢水進流量日統計] TEXT, [廢水處理量] TEXT, [廢水處理量日統計] TEXT, [水站廢水排放量] TEXT, [水站廢水排放量日統計] TEXT, [納管排放量] TEXT, [納管排放量日統計] TEXT, [回收水6吋] TEXT, [回收水6吋日統計] TEXT, [回收水雙介質A] TEXT, [回收水雙介質A日統計] TEXT, [回收水雙介質B] TEXT, [回收水雙介質B日統計] TEXT, [軟水A通量] TEXT, [軟水B通量] TEXT, [軟水C通量] TEXT, [濃縮水至冷卻水池] TEXT, [濃縮水至冷卻水池日統計] TEXT, [濃縮水至逆洗池] TEXT, [濃縮水至逆洗池日統計] TEXT, [貯存池至循環水池] TEXT, [貯存池至循環水池日統計] TEXT, [製程式至循環水池] TEXT, [製程式至循環水池日統計] TEXT, [污泥產出KG] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterChemicals", "[日期] TEXT, [星期] TEXT, [PAC_KG] TEXT, [NAOH_KG] TEXT, [高分子_KG] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterUsageDaily", "[日期] TEXT, [星期] TEXT, [廠區自來水使用量] TEXT, [行政區自來水使用量] TEXT, [自來水至貯存池] TEXT, [自來水至貯存池日統計] TEXT, [自來水量至清水池] TEXT, [自來水量至清水池日統計] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
            // 🟢 月報表強制改為年月
            { "DischargeData", "[年月] TEXT, [水量] TEXT, [SS] TEXT, [COD] TEXT, [BOD] TEXT, [氨氮] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterVolume", "[年月] TEXT, [廠區自來水繳費單] TEXT, [行政區自來水繳費單] TEXT, [彰濱二廠自來水繳費單] TEXT, [附件檔案] TEXT, [備註] TEXT" }
        };

        public App_Water_Generic(string dbName, string tableName, string chineseTitle)
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
            
            // 🟢 自動資料庫升級：將舊的「月份」重新命名為「年月」
            if (columns.Contains("月份")) 
            {
                try 
                {
                    DataManager.RenameColumn(_dbName, _tableName, "月份", "年月");
                    columns = DataManager.GetColumnNames(_dbName, _tableName); 
                } 
                catch { }
            }

            // 🟢 智慧判斷優先順序：日期 > 年月
            if (columns.Contains("日期")) 
            { 
                _timeMode = TimeMode.Date; 
                _dateColumnName = "日期"; 
                if (!columns.Contains("星期")) DataManager.AddColumn(_dbName, _tableName, "星期");
            }
            else if (columns.Contains("年月")) 
            { 
                _timeMode = TimeMode.YearMonth; 
                _dateColumnName = "年月"; 
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
                        _timeMode = TimeMode.Date; 
                        _dateColumnName = columns.FirstOrDefault(c => c != "Id") ?? "Id"; 
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

            if (_timeMode == TimeMode.YearMonth) 
            {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-6));
            }
            else 
            {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            }
            
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

            // 🟢 動態隱藏下拉選單 (針對年月模式隱藏日)
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
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(100, 35) };
            bAdd.Click += async (s, e) => { 
                if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) 
                { 
                    DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); 
                    await LoadGridDataAsync(); 
                    _txtNewColName.Clear(); 
                } 
            };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; 
            _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "修改名稱", Size = new Size(100, 35) };
            bRen.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) 
                { 
                    DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); 
                    await LoadGridDataAsync(); 
                    _txtRenameCol.Clear(); 
                } 
            };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(100, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) 
                { 
                    if(MessageBox.Show($"確定刪除整欄【{_cboColumns.SelectedItem}】？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    { 
                        DataManager.DropColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString()); 
                        await LoadGridDataAsync(); 
                    } 
                } 
            };
            
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += async (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>()
                                       .Select(c => c.OwningRow)
                                       .Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value)
                                       .Distinct().ToList();
                                       
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(包含所屬的實體附件檔案也將被永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                {
                    if (AuthManager.VerifyUser()) 
                    {
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
                            DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                        }
                        await LoadGridDataAsync(); 
                        MessageBox.Show("刪除成功！");
                    }
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bLimitRead.Click += async (s, e) => { 
                if (int.TryParse(txtLimit.Text, out int l)) 
                { 
                    SetUIState(false, "讀取中...", Color.Orange);
                    DataTable dt = null;
                    await Task.Run(() => {
                        dt = DataManager.GetLatestRecords(_dbName, _tableName, l); 
                        EnforceDateFormats(dt);
                    });
                    _dgv.DataSource = dt; 
                    ApplyGridStyles(); 
                    RestoreColumnOrder();
                    SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
                } 
            };
            
            rowAdv2.Controls.AddRange(new Control[] { new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead });
            
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
        // 🟢 附件檔案專用事件與清理機制 (多檔案支援)
        // ==========================================
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
                    
                    using (var frm = new AttachmentForm(currentVal, _dbName, _tableName, path => DeletePhysicalFile(path, e.RowIndex))) 
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
                        if (paths.Contains(relativePath)) 
                        { 
                            isUsedByOthers = true; 
                            break; 
                        }
                    }
                }
            }
            
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
                            dir.Delete(); 
                            dir = dir.Parent;
                        } 
                        else 
                        { 
                            break; 
                        }
                    }
                }
            } 
            catch { }
        }

        private void ApplyGridStyles() 
        {
            if (_dgv.Columns.Contains("Id")) 
            {
                _dgv.Columns["Id"].ReadOnly = true;
            }
            
            if (_dgv.Columns.Contains(_dateColumnName)) 
            {
                string fmt = "yyyy-MM-dd";
                if (_timeMode == TimeMode.YearMonth) fmt = "yyyy-MM";
                
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
        }

        // ==========================================
        // 🟢 資料庫與日期操作邏輯
        // ==========================================
        private void EnforceDateFormats(DataTable dt) 
        {
            if (dt == null || !dt.Columns.Contains(_dateColumnName)) return;
            
            string format = "yyyy-MM-dd";
            if (_timeMode == TimeMode.YearMonth) format = "yyyy-MM";
            
            foreach (DataRow row in dt.Rows) 
            {
                if (row.RowState == DataRowState.Deleted) continue;
                
                string val = row[_dateColumnName]?.ToString();
                
                if (!string.IsNullOrWhiteSpace(val)) 
                {
                    val = val.Replace("/", "-");
                    if (DateTime.TryParse(val, out DateTime d)) 
                    {
                        row[_dateColumnName] = d.ToString(format);
                    }
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

        private string GetDateString(ComboBox y, ComboBox m, ComboBox d) 
        {
            if (_timeMode == TimeMode.YearMonth) return $"{y.SelectedItem}-{m.SelectedItem}";
            
            return $"{y.SelectedItem}-{m.SelectedItem}-{d.SelectedItem}";
        }

        private void UpdateCboColumns() 
        {
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) 
            {
                if (c.Name != "Id" && c.Name != _dateColumnName) 
                {
                    _cboColumns.Items.Add(c.Name);
                }
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) 
        {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            d.SelectedItem = date.Day.ToString("D2");
        }

        private void SaveColumnOrder() 
        { 
            try 
            { 
                var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); 
                File.WriteAllText($"ColOrder_{_dbName}_{_tableName}.txt", string.Join(",", ordered), Encoding.UTF8); 
            } 
            catch { } 
        }
        
        private void RestoreColumnOrder() 
        { 
            try 
            { 
                string fn = $"ColOrder_{_dbName}_{_tableName}.txt"; 
                if (File.Exists(fn)) 
                { 
                    string[] saved = File.ReadAllText(fn, Encoding.UTF8).Split(','); 
                    for (int i = 0; i < saved.Length; i++) 
                    {
                        if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; 
                    }
                } 
            } 
            catch { } 
        }

        private async void BtnSave_Click(object sender, EventArgs e) 
        {
            try 
            {
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

                if (success) 
                {
                    SetUIState(true, "資料儲存成功！", Color.Green);
                    MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    await LoadGridDataAsync();
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
                        SetUIState(false, "Excel 解析與水資源差值運算中，請稍候...", Color.Orange);

                        DataTable originalDt = (DataTable)_dgv.DataSource;
                        _dgv.DataSource = null; 
                        
                        DataTable newBoundDt = null;

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

                                DataTable excelDt = originalDt.Clone();
                                DateTime? minImportDate = null;
                                DateTime? maxImportDate = null;
                                
                                string dateFormat = "yyyy-MM-dd";
                                if (_timeMode == TimeMode.YearMonth) dateFormat = "yyyy-MM";

                                for (int r = 2; r <= rowCount; r++) 
                                {
                                    DataRow nr = excelDt.NewRow();
                                    bool hasData = false;

                                    for (int c = 1; c <= colCount; c++) 
                                    {
                                        string cn = headers[c - 1];
                                        string val = ws.Cells[r, c].Text.Trim(); 

                                        if (excelDt.Columns.Contains(cn) && cn != "Id" && !string.IsNullOrEmpty(val)) 
                                        {
                                            nr[cn] = val;
                                            hasData = true;
                                            
                                            if (cn == _dateColumnName) 
                                            {
                                                string dStr = val.Replace("/", "-");
                                                if (DateTime.TryParse(dStr, out DateTime d)) 
                                                {
                                                    nr[cn] = d.ToString(dateFormat);
                                                    if (minImportDate == null || d < minImportDate) minImportDate = d;
                                                    if (maxImportDate == null || d > maxImportDate) maxImportDate = d;
                                                }
                                            }
                                        }
                                    }
                                    if (hasData) excelDt.Rows.Add(nr);
                                }

                                if (excelDt.Rows.Count == 0) return;

                                DataTable allDbData = DataManager.GetTableData(_dbName, _tableName, "", "", ""); 

                                if (minImportDate.HasValue) 
                                {
                                    string minDateStr = minImportDate.Value.ToString(dateFormat);
                                    
                                    var baselineRow = allDbData.Rows.Cast<DataRow>()
                                        .Where(row => string.Compare(row[_dateColumnName]?.ToString(), minDateStr) < 0)
                                        .OrderByDescending(row => row[_dateColumnName]?.ToString())
                                        .FirstOrDefault();
                                    
                                    if (baselineRow != null) 
                                    {
                                        bool exists = false;
                                        string baseId = baselineRow["Id"].ToString();
                                        foreach(DataRow existing in originalDt.Rows) 
                                        {
                                            if (existing.RowState != DataRowState.Deleted && existing["Id"].ToString() == baseId) 
                                            {
                                                exists = true; 
                                                break;
                                            }
                                        }
                                        if (!exists) originalDt.ImportRow(baselineRow);
                                    }
                                }

                                if (maxImportDate.HasValue) 
                                {
                                    string maxDateStr = maxImportDate.Value.ToString(dateFormat);
                                    
                                    var futureRow = allDbData.Rows.Cast<DataRow>()
                                        .Where(row => string.Compare(row[_dateColumnName]?.ToString(), maxDateStr) > 0)
                                        .OrderBy(row => row[_dateColumnName]?.ToString())
                                        .FirstOrDefault();
                                    
                                    if (futureRow != null) 
                                    {
                                        bool exists = false;
                                        string futId = futureRow["Id"].ToString();
                                        foreach(DataRow existing in originalDt.Rows) 
                                        {
                                            if (existing.RowState != DataRowState.Deleted && existing["Id"].ToString() == futId) 
                                            {
                                                exists = true; 
                                                break;
                                            }
                                        }
                                        if (!exists) originalDt.ImportRow(futureRow);
                                    }
                                }

                                foreach(DataRow r in excelDt.Rows) 
                                {
                                    string importDate = r[_dateColumnName]?.ToString();
                                    if(string.IsNullOrEmpty(importDate)) continue;

                                    DataRow existingRow = originalDt.Rows.Cast<DataRow>()
                                        .FirstOrDefault(row => row.RowState != DataRowState.Deleted && row[_dateColumnName]?.ToString() == importDate);
                                    
                                    if (existingRow != null) 
                                    {
                                        foreach(DataColumn col in excelDt.Columns) 
                                        {
                                            string colName = col.ColumnName;
                                            if (colName == "Id" || colName == _dateColumnName) continue;
                                            
                                            if (r[colName] != DBNull.Value && !string.IsNullOrEmpty(r[colName].ToString())) 
                                            {
                                                if (existingRow[colName].ToString() != r[colName].ToString()) 
                                                {
                                                    existingRow[colName] = r[colName]; 
                                                }
                                            }
                                        }
                                    } 
                                    else 
                                    {
                                        originalDt.ImportRow(r);
                                    }
                                }

                                originalDt.DefaultView.Sort = $"{_dateColumnName} ASC";
                                newBoundDt = originalDt.DefaultView.ToTable(); 

                                _calcHelper?.RecalculateTable(newBoundDt); 
                                _calcHelper?.EndBulkUpdate();
                                EnforceDateFormats(newBoundDt); 
                            }
                        });

                        if (newBoundDt != null) 
                        {
                            _dgv.DataSource = newBoundDt; 
                        } 
                        else 
                        {
                            _dgv.DataSource = originalDt;
                        }
                        
                        ApplyGridStyles();
                        RestoreColumnOrder();

                        SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{((DataTable)_dgv.DataSource).Rows.Count}", Color.Green);
                        MessageBox.Show("Excel 匯入成功！\n系統已自動撈取接軌數據計算差值，並合併重複日期。\n請檢查數據後點擊「儲存數據」。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // ==========================================
        // 🟢 全新四區塊：多檔附件管理視窗
        // ==========================================
        private class AttachmentForm : Form
        {
            public string FinalPathsString { get; private set; }
            private List<string> _paths = new List<string>();
            private string _dbName, _tableName;
            private Action<string> _deleteAction;
            private FlowLayoutPanel _flpList;

            public AttachmentForm(string currentRelPathStr, string dbName, string tableName, Action<string> deleteAction) 
            {
                _dbName = dbName; 
                _tableName = tableName; 
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
                    
                    // 🟢 [修改] 調整按鈕寬度，以容納三個按鈕
                    Button bOpen = new Button { Text = "開啟", Width = 80, Dock = DockStyle.Right, BackColor = Color.LightGray, Cursor = Cursors.Hand };
                    bOpen.Click += (s, e) => { 
                        try { System.Diagnostics.Process.Start(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path)); } 
                        catch (Exception ex) { MessageBox.Show("開啟失敗：" + ex.Message); } 
                    };

                    // 🟢 [新增] 下載/另存新檔按鈕
                    Button bDownload = new Button { Text = "下載", Width = 80, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
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
                    
                    Button bDel = new Button { Text = "刪除", Width = 80, Dock = DockStyle.Right, BackColor = Color.LightCoral, ForeColor = Color.White, Cursor = Cursors.Hand };
                    bDel.Click += (s, e) => { 
                        if (MessageBox.Show($"確定刪除 {Path.GetFileName(path)}?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                        { 
                            _deleteAction(path); 
                            _paths.Remove(path); 
                            RefreshListUI(); 
                        } 
                    };
                    
                    pItem.Controls.Add(lName); 
                    // 🟢 注意加入的順序會影響 Dock.Right 的排列 (越晚加的越靠右)
                    pItem.Controls.Add(bDel);       // 最右邊：刪除
                    pItem.Controls.Add(bDownload);  // 中間：下載
                    pItem.Controls.Add(bOpen);      // 左邊：開啟
                    
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
                
                string datePart = DateTime.Now.ToString("yyyy-MM");
                string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件", _dbName, _tableName, datePart);
                
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
                        
                        File.Copy(src, destPath); 
                        _paths.Add(Path.Combine("附件", _dbName, _tableName, datePart, destName));
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show($"上傳檔案 {Path.GetFileName(src)} 失敗: {ex.Message}", "錯誤"); 
                    }
                }
                RefreshListUI();
            }
        }
    }
}
