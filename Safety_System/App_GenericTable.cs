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
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

        // 🟢 紀錄使用者隱藏欄位的設定
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

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
            string schema = TableSchemaManager.SchemaMap.ContainsKey(_tableName) ? TableSchemaManager.SchemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            LoadVisibilitySettings();

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

            _btnRead = new Button { Text = "🔍 讀取資料", Size = new Size(130, 35), BackColor = Color.WhiteSmoke };
            _btnRead.Click += async (s, e) => { _isFirstLoad = false; await LoadGridDataAsync(); };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(130, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click; 
            
            _btnExport = new Button { Text = "📤 匯出 Excel", Size = new Size(130, 35) }; 
            _btnExport.Click += BtnExport_Click;

            _btnImport = new Button { Text = "📥 匯入 Excel", Size = new Size(130, 35) }; 
            _btnImport.Click += BtnImportExcel_Click;

            _btnExportPdf = new Button { Text = "📄 導出 PDF", Size = new Size(130, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            _btnExportPdf.Click += BtnExportPdf_Click;

            // 🟢 需求3：寬度+20px (160 -> 180)
            _btnColSettings = new Button { Text = "👁️ 欄位顯示設定", Size = new Size(180, 35), BackColor = Color.LightSlateGray, ForeColor = Color.White, Margin = new Padding(15, 0, 0, 0) };
            _btnColSettings.Click += BtnColSettings_Click;

            // 🟢 需求3：寬度+20px (140 -> 160)
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

            // 🟢 需求2：將 _btnExportPdf 與 _btnColSettings 從這行移除
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
                { 
                    DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); 
                    await LoadGridDataAsync(); 
                    _txtNewColName.Clear(); 
                } 
            };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; 
            _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "修改名稱", Size = new Size(120, 35) };
            bRen.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) 
                { 
                    DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); 
                    await LoadGridDataAsync(); 
                    _txtRenameCol.Clear(); 
                } 
            };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
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
            
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(140, 35), BackColor = Color.IndianRed, ForeColor = Color.White, Margin = new Padding(0, 0, 15, 0) };
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

            // 🟢 需求2：將 _btnExportPdf 與 _btnColSettings 加入到 bDelRow 的後方
            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow, _btnExportPdf, _btnColSettings });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0), WrapContents = false };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(140, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
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
            
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            _ = LoadGridDataAsync(); 
            return main;
        }

        // ==========================================
        // 🟢 狀態與控制 UI 輔助方法
        // ==========================================
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

        // ==========================================
        // 🟢 支援按鍵直接輸入 & Alt+Enter 換行
        // ==========================================
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
