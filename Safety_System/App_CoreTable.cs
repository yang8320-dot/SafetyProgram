/// FILE: Safety_System/App_CoreTable.cs ///
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
    public class App_CoreTable
    {
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
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        private readonly ITableLogic _logic; 
        
        private string _dateColumnName = "日期";
        private DataGridViewAutoCalcHelper _calcHelper; 

        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();
        private Dictionary<string, int> _columnWidths = new Dictionary<string, int>();

        // 🟢 右鍵選單與凍結狀態紀錄
        private ContextMenuStrip _ctxMenu;
        private int _rightClickedColIndex = -1;
        private string _frozenColumnName = null;

        public App_CoreTable(string dbName, string tableName, string chineseTitle, ITableLogic logic)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
            _logic = logic ?? new DefaultLogic(); 
        }

        private string GetExpectedFolderName(string rowDateStr)
        {
            if (string.IsNullOrWhiteSpace(rowDateStr)) return DateTime.Now.ToString("yyyy-MM");
            if (_timeMode == TimeMode.Year && rowDateStr.Length >= 4) return rowDateStr.Substring(0, 4);
            if (rowDateStr.Length >= 7) return rowDateStr.Substring(0, 7);
            return DateTime.Now.ToString("yyyy-MM");
        }

        public Control GetView()
        {
            string schema = TableSchemaManager.SchemaMap.ContainsKey(_tableName) 
                            ? TableSchemaManager.SchemaMap[_tableName] 
                            : TableSchemaManager.DefaultCustomSchema;

            DataManager.InitTable(_dbName, _tableName, $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});");
            
            _logic.InitializeSchema(_dbName, _tableName);

            LoadVisibilitySettings();
            LoadColumnWidths(); 

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

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, Padding = new Padding(15) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            Padding lblPad = new Padding(0, 8, 5, 0); 
            Padding ctrlPad = new Padding(0, 4, 5, 0); 
            Padding btnPad = new Padding(0, 0, 10, 0); 
            int btnHeight = 35; 

            // =========================================================
            // 一般顯示區 (1大框)
            // =========================================================
            GroupBox boxTop = new GroupBox { 
                Text = $"{_chineseTitle} (庫：{_dbName} 表：{_tableName})", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                AutoSize = true, 
                AutoSizeMode = AutoSizeMode.GrowAndShrink, 
                Padding = new Padding(10, 15, 10, 10), 
                Margin = new Padding(0, 0, 0, 10) 
            };
            
            FlowLayoutPanel flpTop = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };

            Label lblRange = new Label { Text = "查詢區間:", AutoSize = true, Margin = lblPad };
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };

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

            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, (_timeMode == TimeMode.Date) ? DateTime.Today.AddDays(-30) : DateTime.Today.AddMonths(-6));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            _btnRead = new Button { Text = "🔍 查詢", Size = new Size(90, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            _btnRead.FlatAppearance.BorderSize = 0;
            _btnRead.Click += async (s, e) => { _isFirstLoad = false; _currentSearchMode = SearchMode.DateRange; await ReloadCurrentDataAsync(); };

            Label lblLatestCount = new Label { Text = "最近筆數:", AutoSize = true, Margin = new Padding(15, 8, 5, 0) };
            _txtLatestCount = new TextBox { Width = 50, Text = "50", Margin = ctrlPad }; 
            
            Button bLimitRead = new Button { Text = "查詢", Size = new Size(80, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            bLimitRead.FlatAppearance.BorderSize = 0;
            bLimitRead.Click += async (s, e) => { if (int.TryParse(_txtLatestCount.Text, out int l)) { _isFirstLoad = false; _currentSearchMode = SearchMode.Limit; _currentLimit = l; await ReloadCurrentDataAsync(); } };

            Label lblSY = new Label { Text = "年", AutoSize = true, Margin = lblPad };
            Label lblSM = new Label { Text = "月", AutoSize = true, Margin = lblPad };
            Label lblSD = new Label { Text = "日", AutoSize = true, Margin = lblPad };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            Label lblEY = new Label { Text = "年", AutoSize = true, Margin = lblPad };
            Label lblEM = new Label { Text = "月", AutoSize = true, Margin = lblPad };
            Label lblED = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 15, 0) }; 

            if (_timeMode == TimeMode.YearMonth || _timeMode == TimeMode.Year) {
                _cboStartDay.Visible = false; lblSD.Visible = false;
                _cboEndDay.Visible = false; lblED.Visible = false;
            }
            if (_timeMode == TimeMode.Year) {
                _cboStartMonth.Visible = false; lblSM.Visible = false;
                _cboEndMonth.Visible = false; lblEM.Visible = false;
            }

            _btnToggle = new Button { Text = "+", Size = new Size(45, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(142, 142, 147), ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold) };
            _btnToggle.FlatAppearance.BorderSize = 0;

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存", Size = new Size(95, btnHeight), Margin = new Padding(0, 0, 10, 0), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.FlatAppearance.BorderSize = 0;
            _btnSave.Click += BtnSave_Click; 

            flpTop.Controls.AddRange(new Control[] {
                lblRange, _cboStartYear, lblSY, _cboStartMonth, lblSM, _cboStartDay, lblSD,
                lblTilde, _cboEndYear, lblEY, _cboEndMonth, lblEM, _cboEndDay, lblED,
                _btnRead, lblLatestCount, _txtLatestCount, bLimitRead,
                _btnToggle, _btnSave
            });

            boxTop.Controls.Add(flpTop);

            // =========================================================
            // 進階管理區 (取消分隔，全列接續顯示)
            // =========================================================
            _boxAdvanced = new GroupBox { 
                Text = "進階欄位與權限操作", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F), 
                AutoSize = true, 
                AutoSizeMode = AutoSizeMode.GrowAndShrink, 
                Visible = false, 
                Padding = new Padding(10, 15, 10, 10), 
                ForeColor = Color.DimGray, 
                Margin = new Padding(0, 0, 0, 10) 
            };
            
            TableLayoutPanel tlpAdvLeft = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                ColumnCount = 1, 
                RowCount = 2, 
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink 
            };
            tlpAdvLeft.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpAdvLeft.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            tlpAdvLeft.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            FlowLayoutPanel flpAdvRow1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            FlowLayoutPanel flpAdvRow2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, Margin = new Padding(0, 5, 0, 0) };

            // 🟢 第一排：欄位/列操作→標題更改→刪除欄→刪除列 --> 墊片 --> 顯示設定
            _txtNewColName = new TextBox { Width = 110, Margin = ctrlPad };
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(95, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            bAdd.FlatAppearance.BorderSize = 0;
            bAdd.Click += (s, e) => { 
                if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { 
                    UnfreezeAllColumns();
                    DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); 
                    DataTable dt = (DataTable)_dgv.DataSource;
                    if (!dt.Columns.Contains(_txtNewColName.Text)) dt.Columns.Add(_txtNewColName.Text, typeof(string));
                    ApplyGridStyles(); UpdateCboColumns(); _txtNewColName.Clear(); 
                    ApplyFreezeState();
                    MessageBox.Show("欄位新增成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } 
            };
            
            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad }; 
            _txtRenameCol = new TextBox { Width = 120, Margin = ctrlPad };
            Button bRen = new Button { Text = "標題更改", Size = new Size(95, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            bRen.FlatAppearance.BorderSize = 0;
            bRen.Click += (s, e) => { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { 
                    UnfreezeAllColumns();
                    string oldName = _cboColumns.SelectedItem.ToString();
                    DataManager.RenameColumn(_dbName, _tableName, oldName, _txtRenameCol.Text); 
                    DataTable dt = (DataTable)_dgv.DataSource;
                    if (dt.Columns.Contains(oldName)) dt.Columns[oldName].ColumnName = _txtRenameCol.Text;
                    if (_dgv.Columns.Contains(oldName)) { _dgv.Columns[oldName].HeaderText = _txtRenameCol.Text; _dgv.Columns[oldName].Name = _txtRenameCol.Text; }
                    UpdateCboColumns(); _txtRenameCol.Clear(); 
                    ApplyFreezeState();
                    MessageBox.Show("欄位名稱修改成功！");
                } 
            };
            
            Button bDelCol = new Button { Text = "刪除欄", Size = new Size(80, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 59, 48), ForeColor = Color.White };
            bDelCol.FlatAppearance.BorderSize = 0;
            bDelCol.Click += (s, e) => { 
                if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) { 
                    string colToDrop = _cboColumns.SelectedItem.ToString();
                    if(MessageBox.Show($"確定刪除整欄【{colToDrop}】？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes) { 
                        UnfreezeAllColumns();
                        DataManager.DropColumn(_dbName, _tableName, colToDrop); 
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (dt.Columns.Contains(colToDrop)) dt.Columns.Remove(colToDrop);
                        if (_dgv.Columns.Contains(colToDrop)) _dgv.Columns.Remove(colToDrop);
                        UpdateCboColumns(); 
                        ApplyFreezeState();
                        MessageBox.Show("欄位刪除成功！");
                    } 
                } 
            };
            
            Button bDelRow = new Button { Text = "🗑 刪除列", Size = new Size(110, btnHeight), Margin = new Padding(0,0,0,0), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 59, 48), ForeColor = Color.White };
            bDelRow.FlatAppearance.BorderSize = 0;
            bDelRow.Click += (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>().Select(c => c.OwningRow).Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value).Distinct().ToList();
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    if (AuthManager.VerifyUser()) {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        foreach (var r in selectedRows) {
                            if (_dgv.Columns.Contains("附件檔案")) {
                                string relPathStr = r.Cells["附件檔案"].Value?.ToString();
                                if (!string.IsNullOrEmpty(relPathStr)) {
                                    string[] paths = relPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var p in paths) DeletePhysicalFile(p, r.Index);
                                }
                            }
                            int id = Convert.ToInt32(r.Cells["Id"].Value);
                            DataManager.DeleteRecord(_dbName, _tableName, id);
                            DataRow rowToDelete = dt.Rows.Cast<DataRow>().FirstOrDefault(dr => dr.RowState != DataRowState.Deleted && Convert.ToInt32(dr["Id"]) == id);
                            if (rowToDelete != null) rowToDelete.Delete();
                        }
                        dt.AcceptChanges(); MessageBox.Show("刪除成功！");
                    }
                }
            };

            _btnColSettings = new Button { Text = "👁️ 顯示設定", Size = new Size(120, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(142, 142, 147), ForeColor = Color.White };
            _btnColSettings.FlatAppearance.BorderSize = 0;
            _btnColSettings.Click += BtnColSettings_Click;

            flpAdvRow1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = lblPad }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            flpAdvRow1.Controls.Add(new Panel { Width = 30, Height = 1 }); // 墊片
            flpAdvRow1.Controls.Add(_btnColSettings);

            // 🟢 第二排：查詢資料→關鍵字→查詢 --> 墊片 --> 匯入→匯出→導出PDF
            _cboSearchColumn = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Margin = ctrlPad };
            _txtSearchKeyword = new TextBox { Width = 180, Margin = ctrlPad };
            _btnAdvancedSearch = new Button { Text = "🔍 查詢", Size = new Size(90, btnHeight), Margin = new Padding(0,0,0,0), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            _btnAdvancedSearch.FlatAppearance.BorderSize = 0;
            _btnAdvancedSearch.Click += async (s, e) => { _isFirstLoad = false; _currentSearchMode = SearchMode.Advanced; await ReloadCurrentDataAsync(); };

            _btnImport = new Button { Text = "📥 匯入", Size = new Size(90, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 149, 0), ForeColor = Color.White };
            _btnImport.FlatAppearance.BorderSize = 0;
            _btnImport.Click += BtnImportExcel_Click;

            _btnExport = new Button { Text = "📤 匯出", Size = new Size(90, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White };
            _btnExport.FlatAppearance.BorderSize = 0;
            _btnExport.Click += BtnExport_Click;

            _btnExportPdf = new Button { Text = "📄 導出 PDF", Size = new Size(120, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White };
            _btnExportPdf.FlatAppearance.BorderSize = 0;
            _btnExportPdf.Click += BtnExportPdf_Click;

            flpAdvRow2.Controls.AddRange(new Control[] { new Label { Text = "查詢資料:", AutoSize = true, Margin = lblPad }, _cboSearchColumn, new Label { Text = "關鍵字(含):", AutoSize = true, Margin = lblPad }, _txtSearchKeyword, _btnAdvancedSearch });
            flpAdvRow2.Controls.Add(new Panel { Width = 30, Height = 1 }); // 墊片
            flpAdvRow2.Controls.AddRange(new Control[] { _btnImport, _btnExport, _btnExportPdf });

            if (_logic is LawLogic) {
                _btnRtfToExcel = new Button { Text = "📄 RTF轉 Excel", Size = new Size(160, btnHeight), Margin = btnPad, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(52, 199, 89), ForeColor = Color.White };
                _btnRtfToExcel.FlatAppearance.BorderSize = 0;
                _btnRtfToExcel.Click += BtnRtfToExcel_Click;
                flpAdvRow2.Controls.Add(_btnRtfToExcel);
            }

            tlpAdvLeft.Controls.Add(flpAdvRow1, 0, 0);
            tlpAdvLeft.Controls.Add(flpAdvRow2, 0, 1);
            _boxAdvanced.Controls.Add(tlpAdvLeft);

            _btnToggle.Click += (s, e) => { _boxAdvanced.Visible = !_boxAdvanced.Visible; _btnToggle.Text = _boxAdvanced.Visible ? "-" : "+"; };

            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, AllowUserToResizeColumns = true, 
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells, AllowUserToOrderColumns = true, Margin = new Padding(0, 10, 0, 10)
            };
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _dgv.CellFormatting += Dgv_CellFormatting;
            _dgv.CellClick += Dgv_CellClick;
            _dgv.CellMouseClick += Dgv_CellMouseClick; // 🟢 右鍵選單綁定
            _dgv.KeyDown += Dgv_KeyDown;
            _dgv.KeyPress += Dgv_KeyPress;
            _dgv.EditingControlShowing += Dgv_EditingControlShowing;
            _dgv.DataError += (s, e) => { e.ThrowException = false; };
            _dgv.CurrentCellDirtyStateChanged += Dgv_CurrentCellDirtyStateChanged;
            _dgv.CellValueChanged += Dgv_CellValueChanged;
            _dgv.ColumnWidthChanged += Dgv_ColumnWidthChanged;

            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            InitContextMenu(); // 🟢 初始化右鍵選單

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            _ = ReloadCurrentDataAsync(); 
            return main;
        }

        // ==========================================
        // 右鍵選單與凍結視窗引擎
        // ==========================================
        private void InitContextMenu()
        {
            _ctxMenu = new ContextMenuStrip { Font = new Font("Microsoft JhengHei UI", 11F) };

            ToolStripMenuItem itemFreeze = new ToolStripMenuItem("❄️ 凍結此欄(含)以左視窗");
            ToolStripMenuItem itemUnfreeze = new ToolStripMenuItem("🔥 取消凍結");
            ToolStripMenuItem itemImport = new ToolStripMenuItem("📥 匯入");
            ToolStripMenuItem itemExport = new ToolStripMenuItem("📤 匯出");
            ToolStripMenuItem itemPdf = new ToolStripMenuItem("📄 導出 PDF");

            itemImport.Click += BtnImportExcel_Click;
            itemExport.Click += BtnExport_Click;
            itemPdf.Click += BtnExportPdf_Click;

            itemFreeze.Click += (s, e) => {
                if (_rightClickedColIndex >= 0 && _rightClickedColIndex < _dgv.Columns.Count) {
                    _frozenColumnName = _dgv.Columns[_rightClickedColIndex].Name;
                    ApplyFreezeState();
                }
            };

            itemUnfreeze.Click += (s, e) => {
                _frozenColumnName = null;
                UnfreezeAllColumns();
            };

            _ctxMenu.Items.AddRange(new ToolStripItem[] { itemFreeze, itemUnfreeze, new ToolStripSeparator(), itemImport, itemExport, itemPdf });
        }

        private void Dgv_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) {
                if (e.ColumnIndex >= 0) {
                    _rightClickedColIndex = e.ColumnIndex;
                    if (e.RowIndex >= 0 && e.RowIndex < _dgv.Rows.Count) {
                        _dgv.ClearSelection();
                        _dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                        _dgv.CurrentCell = _dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    }
                    _ctxMenu.Show(Cursor.Position);
                }
            }
        }

        private void UnfreezeAllColumns()
        {
            if (_dgv == null || _dgv.Columns.Count == 0) return;
            foreach (DataGridViewColumn col in _dgv.Columns) 
            {
                col.Frozen = false;
            }
        }

        private void ApplyFreezeState()
        {
            if (string.IsNullOrEmpty(_frozenColumnName) || _dgv == null || !_dgv.Columns.Contains(_frozenColumnName)) return;

            try 
            {
                UnfreezeAllColumns(); 
                int targetIndex = _dgv.Columns[_frozenColumnName].DisplayIndex;
                
                var colsToFreeze = _dgv.Columns.Cast<DataGridViewColumn>()
                                      .Where(c => c.Visible && c.DisplayIndex <= targetIndex)
                                      .OrderBy(c => c.DisplayIndex)
                                      .ToList();
                                      
                foreach(var col in colsToFreeze) {
                    col.Frozen = true;
                }
            } 
            catch { }
        }

        private async Task ReloadCurrentDataAsync()
        {
            int firstRowIndex = -1, selectedRowIndex = -1, selectedColIndex = -1;
            
            if (_dgv.Rows.Count > 0 && !_isFirstLoad) 
            {
                firstRowIndex = _dgv.FirstDisplayedScrollingRowIndex;
                if (_dgv.CurrentCell != null) 
                { 
                    selectedRowIndex = _dgv.CurrentCell.RowIndex; 
                    selectedColIndex = _dgv.CurrentCell.ColumnIndex; 
                }
            }

            if (_currentSearchMode == SearchMode.Advanced) 
            {
                await ExecuteAdvancedSearchAsync();
            }
            else if (_currentSearchMode == SearchMode.Limit) 
            {
                await LoadLimitDataAsync(_currentLimit);
            }
            else 
            {
                await LoadGridDataAsync();
            }

            try 
            {
                if (firstRowIndex >= 0 && firstRowIndex < _dgv.Rows.Count) 
                {
                    _dgv.FirstDisplayedScrollingRowIndex = firstRowIndex;
                }
                
                if (selectedRowIndex >= 0 && selectedRowIndex < _dgv.Rows.Count && selectedColIndex >= 0) 
                {
                    _dgv.ClearSelection();
                    _dgv.CurrentCell = _dgv.Rows[selectedRowIndex].Cells[selectedColIndex];
                    _dgv.Rows[selectedRowIndex].Selected = true;
                }
            } 
            catch { } 
        }

        private async Task LoadGridDataAsync() 
        {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
            string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);

            await Task.Run(() => 
            {
                dt = _isFirstLoad ? DataManager.GetLatestRecords(_dbName, _tableName, 30) : DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate);
                EnforceDateFormats(dt);
            });

            UnfreezeAllColumns(); // 🟢 重新綁定前強制解除凍結，防呆

            _dgv.DataSource = dt;
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();

            ApplyFreezeState(); // 🟢 恢復凍結狀態

            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
        }

        private async Task LoadLimitDataAsync(int limit) 
        {
            SetUIState(false, "讀取中...", Color.Orange);
            DataTable dt = null;
            
            await Task.Run(() => 
            { 
                dt = DataManager.GetLatestRecords(_dbName, _tableName, limit); 
                EnforceDateFormats(dt); 
            });
            
            UnfreezeAllColumns();

            _dgv.DataSource = dt; 
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();

            ApplyFreezeState();

            SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
        }

        private async Task ExecuteAdvancedSearchAsync() 
        {
            SetUIState(false, "條件搜尋中，請稍候...", Color.Orange);
            string searchCol = _cboSearchColumn.SelectedItem?.ToString();
            string keyword = _txtSearchKeyword.Text;
            DataTable resultDt = null;

            await Task.Run(() => 
            {
                DataTable allData = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                DataView dv = allData.DefaultView;

                if (!string.IsNullOrEmpty(searchCol)) 
                {
                    if (keyword == "有鍵入資料者") 
                        dv.RowFilter = $"[{searchCol}] <> '' AND [{searchCol}] IS NOT NULL";
                    else if (string.IsNullOrWhiteSpace(keyword)) 
                        dv.RowFilter = $"[{searchCol}] IS NULL OR [{searchCol}] = ''";
                    else 
                        dv.RowFilter = $"[{searchCol}] LIKE '%{keyword.Replace("'", "''")}%'";
                }
                
                dv.Sort = "Id DESC"; 
                resultDt = dv.ToTable(); 
                
                if (_logic is LawLogic && int.TryParse(_txtLatestCount.Text, out int limit)) 
                {
                    DataTable limitedDt = resultDt.Clone();
                    for (int i = 0; i < Math.Min(limit, resultDt.Rows.Count); i++) 
                    {
                        limitedDt.ImportRow(resultDt.Rows[i]);
                    }
                    resultDt = limitedDt;
                }
                EnforceDateFormats(resultDt);
            });

            UnfreezeAllColumns();

            _dgv.DataSource = resultDt;
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();

            ApplyFreezeState();

            SetUIState(true, $"搜尋完成，共找到 {resultDt.Rows.Count} 筆資料", Color.Green);
        }

        private void LoadColumnWidths() 
        {
            _columnWidths.Clear();
            var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Width");
            foreach (var kvp in dict) 
            {
                if (int.TryParse(kvp.Value, out int w)) _columnWidths[kvp.Key] = w; 
            }
        }

        private void SaveColumnWidths() 
        {
            foreach (var kvp in _columnWidths) 
            {
                DataManager.SaveGridConfig(_dbName, _tableName, "Width", kvp.Key, kvp.Value.ToString());
            }
        }

        private void Dgv_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e) 
        {
            if (_isFirstLoad || _isApplyingWidths) return;
            if (e.Column != null) 
            { 
                _columnWidths[e.Column.Name] = e.Column.Width; 
                SaveColumnWidths(); 
            }
        }

        private void LoadVisibilitySettings() 
        {
            _columnVisibility.Clear();
            var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Visibility");
            foreach (var kvp in dict) 
            {
                _columnVisibility[kvp.Key] = (kvp.Value == "1");
            }
        }

        private void SaveVisibilitySettings() 
        {
            DataManager.ClearGridConfig(_dbName, _tableName, "Visibility");
            foreach (var kvp in _columnVisibility) 
            {
                DataManager.SaveGridConfig(_dbName, _tableName, "Visibility", kvp.Key, kvp.Value ? "1" : "0");
            }
        }

        private void SaveColumnOrder() 
        { 
            try 
            { 
                var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); 
                DataManager.SaveGridConfig(_dbName, _tableName, "Order", "All", string.Join(",", ordered)); 
            } 
            catch { } 
        }
        
        private void RestoreColumnOrder() 
        { 
            try 
            { 
                UnfreezeAllColumns(); // 🟢 排版重組前解鎖

                var dict = DataManager.LoadGridConfig(_dbName, _tableName, "Order"); 
                if (dict.ContainsKey("All")) 
                { 
                    string[] saved = dict["All"].Split(','); 
                    for (int i = 0; i < saved.Length; i++) 
                    { 
                        if (_dgv.Columns.Contains(saved[i])) 
                            _dgv.Columns[saved[i]].DisplayIndex = i; 
                    } 
                } 
            } 
            catch { } 
        }

        private void BtnColSettings_Click(object sender, EventArgs e) 
        {
            if (_dgv.Columns.Count == 0) return;
            using (Form f = new Form { Text = "👁️ 欄位顯示設定", Size = new Size(350, 500), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false }) 
            {
                Label lblTop = new Label { Text = "請勾選欲顯示在表格中的欄位：", Dock = DockStyle.Top, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.SteelBlue }; 
                f.Controls.Add(lblTop);
                
                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.None, Padding = new Padding(10) };
                
                foreach (DataGridViewColumn col in _dgv.Columns) 
                { 
                    if (col.Name == "Id") continue; 
                    bool isChecked = _columnVisibility.ContainsKey(col.Name) ? _columnVisibility[col.Name] : true; 
                    clbCols.Items.Add(col.Name, isChecked); 
                }
                
                f.Controls.Add(clbCols);
                
                Button btnSave = new Button { Text = "💾 儲存並套用設定", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnSave.Click += (s, ev) => 
                { 
                    for (int i = 0; i < clbCols.Items.Count; i++) 
                    { 
                        string colName = clbCols.Items[i].ToString(); 
                        bool isChecked = clbCols.GetItemChecked(i); 
                        _columnVisibility[colName] = isChecked; 
                        if (_dgv.Columns.Contains(colName)) _dgv.Columns[colName].Visible = isChecked; 
                    } 
                    SaveVisibilitySettings(); 
                    f.DialogResult = DialogResult.OK; 
                };
                
                f.Controls.Add(btnSave); 
                f.ShowDialog();

                ApplyFreezeState(); // 重新觸發凍結狀態，確保隱藏欄位時不會打亂
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
            if(_btnRtfToExcel != null) _btnRtfToExcel.Enabled = isEnabled;
            
            _lblStatus.Text = statusText; 
            _lblStatus.ForeColor = statusColor;
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
                if (_columnVisibility.ContainsKey(col.Name)) col.Visible = _columnVisibility[col.Name];

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

            if (_logic is LawLogic) 
            {
                if (_dgv.Columns.Contains("法規名稱")) _dgv.Columns["法規名稱"].Width = 250;
                if (_dgv.Columns.Contains("內容")) _dgv.Columns["內容"].Width = 400;
                if (_dgv.Columns.Contains("重點摘要")) _dgv.Columns["重點摘要"].Width = 200;
            }

            foreach (DataGridViewColumn col in _dgv.Columns) 
            {
                if (_columnWidths.ContainsKey(col.Name) && _columnWidths[col.Name] > 0)
                    col.Width = _columnWidths[col.Name];
            }

            _isApplyingWidths = false; 
        }

        private void SetupDropdownColumns() 
        {
            foreach (DataGridViewColumn col in _dgv.Columns.Cast<DataGridViewColumn>().ToList()) 
            {
                string[] items = _logic.GetDropdownList(_tableName, col.Name);
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
                            if (!string.IsNullOrEmpty(val) && !finalItems.Contains(val)) finalItems.Add(val); 
                        } 
                    }
                    
                    cboCol.Items.AddRange(finalItems.ToArray()); 
                    _dgv.Columns.Insert(colIndex, cboCol);
                }
            }
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
            // 🟢 徹底攔截控制項按鍵，保障 Ctrl+S 在編輯模式也能用
            e.Control.PreviewKeyDown -= EditingControl_PreviewKeyDown;
            e.Control.PreviewKeyDown += EditingControl_PreviewKeyDown;

            if (e.Control is ComboBox cbo) 
            {
                cbo.DropDownStyle = ComboBoxStyle.DropDownList;
                if (_dgv.CurrentCell != null) 
                {
                    string colName = _dgv.Columns[_dgv.CurrentCell.ColumnIndex].Name;
                    string[] items = null;
                    if (colName == "危害類型細分類") 
                    { 
                        string parentVal = _dgv.CurrentRow.Cells["危害類型主項"].Value?.ToString() ?? ""; 
                        items = _logic.GetDependentDropdownList(_tableName, colName, parentVal); 
                    } 
                    else if (colName == "違規樣態類型") 
                    { 
                        string parentVal = _dgv.CurrentRow.Cells["危害類型細分類"].Value?.ToString() ?? ""; 
                        items = _logic.GetDependentDropdownList(_tableName, colName, parentVal); 
                    }
                    
                    if (items != null) 
                    { 
                        object currentVal = _dgv.CurrentCell.Value; 
                        cbo.Items.Clear(); 
                        cbo.Items.AddRange(items); 
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

        private void EditingControl_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.S)
            {
                e.IsInputKey = true; // 放行讓上層抓取
                _btnSave.PerformClick();
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

        // 🟢 修復下拉選單直接點擊儲存時未刷入緩衝區的 Bug
        private async void BtnSave_Click(object sender, EventArgs e) 
        {
            try 
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                // 強制轉移焦點，迫使編輯元件觸發 Leave 並寫入值
                _btnSave.Focus();

                if (_dgv.IsCurrentCellInEditMode) {
                    _dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
                _dgv.EndEdit(); 
                
                if (_dgv.BindingContext != null && _dgv.DataSource != null) {
                    _dgv.BindingContext[_dgv.DataSource].EndCurrentEdit();
                }

                SaveColumnOrder(); 
                SetUIState(false, "資料庫寫入與檔案同步中，請稍候...", Color.Orange);
                
                DataTable dt = ((DataTable)_dgv.DataSource).Copy(); 
                bool success = false;
                
                await Task.Run(async () => 
                { 
                    EnforceDateFormats(dt); 
                    SyncAttachmentPaths(dt);
                    if (await _logic.OnBeforeSaveAsync(_dbName, _tableName, dt)) 
                    {
                        success = DataManager.BulkSaveTable(_dbName, _tableName, dt); 
                        if (success) await _logic.OnAfterSaveAsync(_dbName, _tableName, dt);
                    }
                });
                
                if (success) 
                { 
                    SetUIState(true, "資料儲存成功！", Color.Green); 
                    MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
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
                else 
                { 
                    _logic.OnCellClick(_dgv, e); 
                }
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
                        foreach(DataRow r in dt.Rows) 
                        { 
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
            
            if (!isUsedByOthers) 
            {
                try 
                { 
                    DataTable dt = DataManager.GetTableData(_dbName, _tableName, "", "", ""); 
                    foreach (DataRow row in dt.Rows) 
                    { 
                        string val = row["附件檔案"]?.ToString(); 
                        if (!string.IsNullOrEmpty(val) && val.Contains(relativePath)) { isUsedByOthers = true; break; } 
                    } 
                } 
                catch { } 
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
                        
                        UnfreezeAllColumns(); // 🟢 匯入重綁定前解開凍結
                        _dgv.DataSource = null; 
                        
                        await Task.Run(() => 
                        {
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
                        ApplyFreezeState();

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

        private void BtnExportPdf_Click(object sender, EventArgs e) 
        {
            if (_dgv.Rows.Count <= 1) { MessageBox.Show("目前沒有資料可供導出。"); return; }
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
                    
                    pd.PrintPage += (s, ev) => 
                    {
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
                        
                        g.DrawString("台灣玻璃工業股份有限公司-彰濱廠", fTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 40), sfCenter); 
                        y += 35;
                        g.DrawString(_chineseTitle, fSubTitle, Brushes.Black, new RectangleF(x, y, pageWidth, 30), sfCenter); 
                        y += 30;
                        
                        string filterStr = ""; 
                        if (!string.IsNullOrEmpty(_txtSearchKeyword.Text)) filterStr = $" | 關鍵字: {_txtSearchKeyword.Text}";
                        g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}{filterStr}", fBody, Brushes.Gray, new RectangleF(x, y, pageWidth, 25), sfLeft); 
                        y += 25;
                        
                        var visCols = _dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList(); 
                        if (visCols.Count == 0) return;
                        
                        float totalGridWidth = visCols.Sum(c => c.Width); 
                        float[] actualColWidths = new float[visCols.Count];
                        for (int i = 0; i < visCols.Count; i++) 
                        { 
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
                            for (int i = 0; i < visCols.Count; i++) 
                            { 
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
                    
                    try 
                    { 
                        pd.Print(); 
                        if (activeForm != null) activeForm.Cursor = Cursors.Default; 
                        MessageBox.Show("PDF 報表匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    } 
                    catch (Exception ex) 
                    { 
                        if (activeForm != null) activeForm.Cursor = Cursors.Default; 
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    } 
                    finally 
                    { 
                        if (activeForm != null) activeForm.Cursor = Cursors.Default; 
                    }
                }
            }
        }

        private void BtnRtfToExcel_Click(object sender, EventArgs e) 
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "RTF 法規檔案 (*.rtf)|*.rtf", Title = "請選擇全國法規資料庫下載的 RTF 檔案" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + "_轉換.xlsx" }) 
                    {
                        if (sfd.ShowDialog() == DialogResult.OK) 
                        {
                            try 
                            { 
                                LawRtfToExcelConverter.Convert(ofd.FileName, sfd.FileName); 
                                MessageBox.Show("轉換成功！\n您現在可以點擊「匯入 EXCEL」將產生的檔案載入系統。", "轉換完成", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                            } 
                            catch (Exception ex) 
                            { 
                                MessageBox.Show("轉換失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                            }
                        }
                    }
                }
            }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e) 
        {
            if (e.Control && e.KeyCode == Keys.S) 
            { 
                e.Handled = true; 
                e.SuppressKeyPress = true; 
                _btnSave?.PerformClick(); 
            }
            else if (e.Control && e.KeyCode == Keys.V) 
            {
                try 
                {
                    string text = Clipboard.GetText(); 
                    if (string.IsNullOrEmpty(text)) return;
                    
                    _calcHelper?.BeginBulkUpdate(); 
                    _dgv.SuspendLayout(); 
                    
                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex; 
                    DataTable dt = (DataTable)_dgv.DataSource;
                    
                    foreach (string line in lines) 
                    {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++) 
                        { 
                            if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly) 
                            { 
                                _dgv[c + i, r].Value = cells[i].Trim().Trim('"'); 
                            } 
                        }
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt); 
                    _calcHelper?.EndBulkUpdate(); 
                    EnforceDateFormats(dt); 
                    _dgv.ResumeLayout();
                } 
                catch 
                { 
                    _calcHelper?.EndBulkUpdate(); 
                    _dgv.ResumeLayout(); 
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

        private void Dgv_CurrentCellDirtyStateChanged(object sender, EventArgs e) 
        {
            if (_dgv.IsCurrentCellDirty && _dgv.CurrentCell is DataGridViewComboBoxCell) 
            { 
                _dgv.CommitEdit(DataGridViewDataErrorContexts.Commit); 
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
                pnlDrop.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
                pnlDrop.DragDrop += (s, e) => { ProcessUpload((string[])e.Data.GetData(DataFormats.FileDrop)); };
                boxUpload.Controls.Add(pnlDrop); 
                tlp.Controls.Add(boxUpload, 0, 1);

                Button btnClearAll = new Button { Text = "🗑️ 清除此筆紀錄的所有附件", Dock = DockStyle.Fill, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(3, 5, 3, 5) };
                btnClearAll.Click += (s, e) => 
                {
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
                btnSaveClose.Click += (s, e) => 
                { 
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
                    bOpen.Click += (s, e) => 
                    { 
                        try { System.Diagnostics.Process.Start(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path)); } 
                        catch (Exception ex) { MessageBox.Show("開啟失敗：" + ex.Message); } 
                    };
                    
                    Button bDownload = new Button { Text = "下載", Width = 100, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
                    bDownload.Click += (s, e) => 
                    {
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
                    bDel.Click += (s, e) => 
                    { 
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
                _encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L); 
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
