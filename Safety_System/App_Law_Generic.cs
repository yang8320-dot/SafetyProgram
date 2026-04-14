/// FILE: Safety_System/App_Law_Generic.cs ///
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
    public class App_Law_Generic
    {
        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;

        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        
        // 將按鈕提至全域以利狀態控制
        private Button _btnToggle;     
        private Button _btnSave; 
        private Button _btnRead;
        private Button _btnExport;
        private Button _btnImport;
        private Button _btnAdvancedSearch;
        private Button _btnRtfToExcel;

        // 🟢 新增：UI 狀態提示列
        private Label _lblStatus;

        private bool _isFirstLoad = true;
        
        private readonly string _dbName; 
        private readonly string _tableName; 

        private ComboBox _cboSearchColumn;
        private TextBox _txtSearchKeyword;
        private TextBox _txtLatestCount;

        private const string DirectoryTableName = "法規目錄一覽";

        public App_Law_Generic(string dbName, string tableName)
        {
            _dbName = dbName;
            _tableName = tableName;
        }

        public Control GetView()
        {
            DataManager.InitTable(_dbName, _tableName, $@"CREATE TABLE IF NOT EXISTS [{_tableName}] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [日期] TEXT, 
                [法規名稱] TEXT, 
                [條] TEXT,
                [項] TEXT,
                [款] TEXT,
                [目] TEXT,
                [內容] TEXT,
                [重點摘要] TEXT, 
                [適用性] TEXT, 
                [有提升績效機會] TEXT,
                [有潛在不符合風險] TEXT,
                [鑑別日期] TEXT,
                [備註] TEXT);");

            DataManager.InitTable(_dbName, DirectoryTableName, $@"CREATE TABLE IF NOT EXISTS [{DirectoryTableName}] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                [選項類別] TEXT, 
                [流水號] TEXT, 
                [法規名稱] TEXT,
                [日期] TEXT,
                [適用性] TEXT,
                [鑑別日期] TEXT,
                [再次確認日期] TEXT);");

            var existingCols = DataManager.GetColumnNames(_dbName, _tableName);
            if (!existingCols.Contains("有提升績效機會")) DataManager.AddColumn(_dbName, _tableName, "有提升績效機會");
            if (!existingCols.Contains("有潛在不符合風險")) DataManager.AddColumn(_dbName, _tableName, "有潛在不符合風險");

            // 🟢 優化 1：TableLayoutPanel 加入 Padding，完美排版
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, Padding = new Padding(15) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));      
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));      
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 給狀態列      
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            // 🟢 優化 2：Margin 下推 10px 隔開距離
            GroupBox boxTop = new GroupBox { Text = $"法規管理 (庫：{_dbName} 表：{_tableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10), Margin = new Padding(0, 0, 0, 10) };
            
            FlowLayoutPanel flpTopContainer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };

            FlowLayoutPanel row1 = new FlowLayoutPanel { AutoSize = true, WrapContents = true };
            Label lblRange = new Label { Text = "法規日期:", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };

            int currentYear = DateTime.Now.Year;
            for (int i = currentYear - 25; i <= currentYear; i++) {
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

            SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddYears(-1));
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            _btnRead = new Button { Text = "區間讀取", Size = new Size(100, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            _btnRead.Click += async (s, e) => { _isFirstLoad = false; await LoadGridDataAsync(); };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Margin = new Padding(5, 0, 0, 0) };
            _btnSave.Click += BtnSave_Click; 
            
            _btnExport = new Button { Text = "匯出 Excel", Size = new Size(100, 35) }; 
            _btnExport.Click += BtnExport_Click;

            _btnToggle = new Button { Text = "[ + ] 欄位管理", Size = new Size(150, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };

            row1.Controls.AddRange(new Control[] { 
                lblRange, 
                _cboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) },
                _cboStartMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) },
                _cboStartDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 8, 5, 0) }, 
                _cboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) },
                _cboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) },
                _cboEndDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 15, 0) },
                _btnRead, _btnExport, _btnToggle, _btnSave 
            });

            FlowLayoutPanel row3 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 5), Font = new Font("Microsoft JhengHei UI", 11F) };
            
            _txtLatestCount = new TextBox { Width = 60, Text = "500", TextAlign = HorizontalAlignment.Center }; 
            _cboSearchColumn = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            _txtSearchKeyword = new TextBox { Width = 180 }; 

            _cboSearchColumn.SelectedIndexChanged += (s, e) => {
                string sel = _cboSearchColumn.SelectedItem?.ToString();
                if (sel == "重點摘要" || sel == "備註") {
                    _txtSearchKeyword.Text = "有鍵入資料者";
                } else if (_txtSearchKeyword.Text == "有鍵入資料者") {
                    _txtSearchKeyword.Text = "";
                }
            };

            _btnAdvancedSearch = new Button { Text = "🔍 條件搜尋", Size = new Size(130, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            _btnAdvancedSearch.Click += async (s, e) => await ExecuteAdvancedSearchAsync();

            _btnRtfToExcel = new Button { Text = "📄 全國法規 RTF 轉 EXCEL", Size = new Size(240, 35), BackColor = Color.DarkSeaGreen, ForeColor = Color.White, Margin = new Padding(15, 0, 0, 0) };
            _btnRtfToExcel.Click += BtnRtfToExcel_Click;

            _btnImport = new Button { Text = "📥 匯入 EXCEL", Size = new Size(140, 35), BackColor = Color.WhiteSmoke, Margin = new Padding(15, 0, 0, 0) }; 
            _btnImport.Click += BtnImportExcel_Click;

            row3.Controls.AddRange(new Control[] { 
                new Label { Text = "顯示最新筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtLatestCount, 
                new Label { Text = "查詢欄位:", AutoSize = true, Margin = new Padding(15, 8, 0, 0) }, _cboSearchColumn, 
                new Label { Text = "關鍵字(包含):", AutoSize = true, Margin = new Padding(15, 8, 0, 0) }, _txtSearchKeyword, 
                _btnAdvancedSearch, _btnRtfToExcel, _btnImport 
            });

            flpTopContainer.Controls.Add(row1);
            flpTopContainer.Controls.Add(row3);
            boxTop.Controls.Add(flpTopContainer);

            // ================= 可隱藏的欄位操作排 =================
            GroupBox boxOps = new GroupBox { Text = "進階欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 10) };
            FlowLayoutPanel row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            _txtNewColName = new TextBox { Width = 120 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(120, 35) }; 
            bAdd.Click += async (s, e) => { if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) { DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); await LoadGridDataAsync(); _txtNewColName.Clear(); } };
            
            _cboColumns = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList }; _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "標題更改", Size = new Size(120, 35) }; 
            bRen.Click += async (s, e) => { if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) { DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); await LoadGridDataAsync(); _txtRenameCol.Clear(); } };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(120, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += async (s, e) => { if (_cboColumns.SelectedItem != null) { string colToDrop = _cboColumns.SelectedItem.ToString(); if (MessageBox.Show(Form.ActiveForm, $"警告：確定要刪除整欄【{colToDrop}】嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) { if (AuthManager.VerifyAdmin()) { DataManager.DropColumn(_dbName, _tableName, colToDrop); await LoadGridDataAsync(); } } } };
            
            Button bDelRow = new Button { Text = "🗑️ 刪除選取列", Size = new Size(140, 35), BackColor = Color.IndianRed, ForeColor = Color.White }; 
            bDelRow.Click += async (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>()
                                       .Select(c => c.OwningRow)
                                       .Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value)
                                       .Distinct().ToList();

                if (selectedRows.Count > 0) {
                    if (MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(刪除後將立即生效)", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        if (AuthManager.VerifyUser()) {
                            foreach (var r in selectedRows) DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                            await LoadGridDataAsync();
                            MessageBox.Show("刪除成功！");
                        }
                    }
                } else {
                    MessageBox.Show("請先用滑鼠選取要刪除的資料列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            
            row2.Controls.AddRange(new Control[] { new Label { Text = "欄位操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            boxOps.Controls.Add(row2);

            _btnToggle.Click += (s, e) => { 
                boxOps.Visible = !boxOps.Visible; 
                _btnToggle.Text = boxOps.Visible ? "[ - ] 隱藏管理" : "[ + ] 欄位管理"; 
                _btnToggle.BackColor = boxOps.Visible ? Color.LightCoral : Color.LightGray; 
            };

            // 🟢 新增：UI 防呆狀態列
            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            // 🟢 優化 3：RowTemplate 增高、增加交替列顏色，提升閱讀體驗
            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                Margin = new Padding(0, 10, 0, 10)
            };
            _dgv.RowTemplate.Height = 35;
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _dgv.DataError += Dgv_DataError;
            _dgv.EditingControlShowing += Dgv_EditingControlShowing;
            _dgv.KeyDown += Dgv_KeyDown; 
            _dgv.CellClick += Dgv_CellClick;

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(boxOps, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);
            
            // 🟢 啟動時非同步載入
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
            _btnAdvancedSearch.Enabled = isEnabled;
            _btnRtfToExcel.Enabled = isEnabled;
            
            _lblStatus.Text = statusText;
            _lblStatus.ForeColor = statusColor;
        }

        private void EnforceDateFormats(DataTable dt)
        {
            if (dt == null) return;
            string[] dateCols = { "日期", "鑑別日期" };
            foreach (DataRow row in dt.Rows) {
                if (row.RowState == DataRowState.Deleted) continue;
                foreach (string col in dateCols) {
                    if (dt.Columns.Contains(col)) {
                        string val = row[col]?.ToString();
                        if (!string.IsNullOrWhiteSpace(val)) {
                            val = val.Replace("/", "-");
                            if (DateTime.TryParse(val, out DateTime d)) {
                                row[col] = d.ToString("yyyy-MM-dd");
                            }
                        }
                    }
                }
            }
        }

        // ==========================================
        // 🟢 核心資料載入 (加入非同步優化)
        // ==========================================
        private async Task LoadGridDataAsync() {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            
            string sDate = GetStartDate().ToString("yyyy-MM-dd");
            string eDate = GetEndDate().ToString("yyyy-MM-dd");

            await Task.Run(() => {
                if (_isFirstLoad) {
                    dt = DataManager.GetLatestRecords(_dbName, _tableName, 0); 
                } else { 
                    dt = DataManager.GetTableData(_dbName, _tableName, "日期", sDate, eDate); 
                }
                EnforceDateFormats(dt);
            });

            _dgv.DataSource = dt;
            ApplyGridFormatting();
            _isFirstLoad = false; 
            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
        }

        private async Task ExecuteAdvancedSearchAsync()
        {
            SetUIState(false, "條件搜尋中，請稍候...", Color.Orange);
            
            string countText = _txtLatestCount.Text;
            string searchCol = _cboSearchColumn.SelectedItem?.ToString();
            string keyword = _txtSearchKeyword.Text;

            int limit = 500;
            if (!int.TryParse(countText, out limit) || limit <= 0) limit = 500; 

            DataTable resultDt = null;

            await Task.Run(() => {
                DataTable allData = DataManager.GetTableData(_dbName, _tableName, "日期", "", "");
                DataView dv = allData.DefaultView;

                if (!string.IsNullOrEmpty(searchCol)) 
                {
                    if (keyword == "有鍵入資料者") {
                        dv.RowFilter = $"[{searchCol}] <> '' AND [{searchCol}] IS NOT NULL";
                    }
                    else if ((searchCol == "適用性" || searchCol == "鑑別日期") && string.IsNullOrWhiteSpace(keyword)) {
                        dv.RowFilter = $"[{searchCol}] IS NULL OR [{searchCol}] = ''";
                    }
                    else if (!string.IsNullOrWhiteSpace(keyword)) {
                        dv.RowFilter = $"[{searchCol}] LIKE '%{keyword.Replace("'", "''")}%'";
                    }
                }
                dv.Sort = "Id DESC"; 
                
                resultDt = dv.ToTable().Clone(); 
                int count = 0;
                foreach (DataRowView drv in dv) {
                    if (count >= limit) break;
                    resultDt.ImportRow(drv.Row);
                    count++;
                }
                EnforceDateFormats(resultDt);
            });

            _dgv.DataSource = resultDt;
            ApplyGridFormatting();
            SetUIState(true, $"搜尋完成，共找到 {resultDt.Rows.Count} 筆符合條件資料", Color.Green);
        }

        private void ApplyGridFormatting()
        {
            SetupComboBoxColumns();
            SetupTextWrapping();

            if (_dgv.Columns.Contains("Id")) {
                _dgv.Columns["Id"].ReadOnly = true;
                _dgv.Columns["Id"].Visible = false;
            }
            if (_dgv.Columns.Contains("日期")) _dgv.Columns["日期"].DefaultCellStyle.Format = "yyyy-MM-dd";
            if (_dgv.Columns.Contains("鑑別日期")) _dgv.Columns["鑑別日期"].DefaultCellStyle.Format = "yyyy-MM-dd";

            UpdateCboColumns();
            _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        }

        // ==========================================
        // 🟢 核心儲存邏輯 (完整保留原有法規目錄與重複比對)
        // ==========================================
        private async void BtnSave_Click(object sender, EventArgs e)
        {
            try {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit(); 
                SetUIState(false, "資料庫運算與儲存中，請稍候...", Color.Orange);
                
                DataTable dtToSave = ((DataTable)_dgv.DataSource).Copy();

                bool success = await Task.Run(() =>
                {
                    EnforceDateFormats(dtToSave);
                    ResolveDuplicates(dtToSave); // 處理重複判斷
                    
                    bool mainSave = DataManager.BulkSaveTable(_dbName, _tableName, dtToSave);
                    if (mainSave) {
                        GenerateAndSaveDirectory(); // 自動重新產生目錄
                    }
                    return mainSave;
                });

                if (success) {
                    SetUIState(true, "儲存完成！(含法規目錄一覽表同步更新)", Color.Green);
                    MessageBox.Show(Form.ActiveForm, "儲存/更新 完成！(已啟用 Transaction 交易機制)\n\n✅ 背景執行寫入目錄一覽表完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    await LoadGridDataAsync(); 
                } else {
                    SetUIState(true, "儲存失敗", Color.Red);
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

        private void GenerateAndSaveDirectory()
        {
            DataTable dtMain = DataManager.GetTableData(_dbName, _tableName, "", "", "");
            DataTable dtDirExist = DataManager.GetTableData(_dbName, DirectoryTableName, "", "", "");
            var existingDirDict = new Dictionary<string, DataRow>();
            
            foreach (DataRow r in dtDirExist.Rows) {
                if (r["選項類別"]?.ToString() == _tableName) {
                    string name = r["法規名稱"]?.ToString().Trim();
                    if (!string.IsNullOrEmpty(name)) existingDirDict[name] = r;
                }
            }

            var grouped = new Dictionary<string, List<DataRow>>();
            foreach(DataRow r in dtMain.Rows) {
                string name = r["法規名稱"]?.ToString().Trim();
                if (string.IsNullOrEmpty(name)) continue;
                
                if (!grouped.ContainsKey(name)) grouped[name] = new List<DataRow>();
                grouped[name].Add(r);
            }

            DataTable dtDir = new DataTable();
            dtDir.Columns.Add("Id", typeof(int)); 
            dtDir.Columns.Add("選項類別", typeof(string));
            dtDir.Columns.Add("流水號", typeof(string));
            dtDir.Columns.Add("法規名稱", typeof(string));
            dtDir.Columns.Add("日期", typeof(string));
            dtDir.Columns.Add("適用性", typeof(string));
            dtDir.Columns.Add("鑑別日期", typeof(string));
            dtDir.Columns.Add("再次確認日期", typeof(string));

            int index = 1;
            HashSet<string> processedNames = new HashSet<string>();

            foreach(var kvp in grouped) {
                string lawName = kvp.Key;
                processedNames.Add(lawName);

                string latestDate = "";
                string latestIdenDate = "";
                string applyStatus = "";
                bool hasApplicable = false;

                foreach(var row in kvp.Value) {
                    string d = row["日期"]?.ToString() ?? "";
                    string iden = row["鑑別日期"]?.ToString() ?? "";
                    string apply = row["適用性"]?.ToString() ?? "";

                    if (string.Compare(d, latestDate) > 0) latestDate = d;
                    if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;

                    if (apply == "適用") hasApplicable = true;
                    if (string.IsNullOrEmpty(applyStatus)) applyStatus = apply; 
                }

                DataRow newRow = dtDir.NewRow();
                newRow["選項類別"] = _tableName; 
                newRow["流水號"] = index.ToString(); 
                newRow["法規名稱"] = lawName;
                newRow["日期"] = latestDate;
                newRow["適用性"] = hasApplicable ? "適用" : applyStatus; 
                newRow["鑑別日期"] = latestIdenDate;

                if (existingDirDict.ContainsKey(lawName)) {
                    newRow["Id"] = existingDirDict[lawName]["Id"];
                    newRow["再次確認日期"] = existingDirDict[lawName]["再次確認日期"]?.ToString();
                } else {
                    newRow["再次確認日期"] = ""; 
                }

                dtDir.Rows.Add(newRow);
                index++;
            }

            foreach (var kvp in existingDirDict) {
                if (!processedNames.Contains(kvp.Key)) {
                    int idToDelete = Convert.ToInt32(kvp.Value["Id"]);
                    DataManager.DeleteRecord(_dbName, DirectoryTableName, idToDelete);
                }
            }

            DataManager.BulkSaveTable(_dbName, DirectoryTableName, dtDir);
        }

        private void ResolveDuplicates(DataTable dt)
        {
            DataTable dbData = DataManager.GetTableData(_dbName, _tableName, "", "", "");
            var existingDict = new Dictionary<string, int>();

            foreach (DataRow dbRow in dbData.Rows)
            {
                string name = dbRow["法規名稱"]?.ToString().Trim() ?? "";
                string article = dbRow["條"]?.ToString().Trim() ?? "";
                string content = dbRow["內容"]?.ToString().Trim() ?? "";
                string key = $"{name}_|{article}_|{content}";
                existingDict[key] = Convert.ToInt32(dbRow["Id"]);
            }

            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                if (row.Table.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0)
                    continue;

                string name = row["法規名稱"]?.ToString().Trim() ?? "";
                string article = row["條"]?.ToString().Trim() ?? "";
                string content = row["內容"]?.ToString().Trim() ?? "";
                string key = $"{name}_|{article}_|{content}";

                if (existingDict.ContainsKey(key))
                {
                    bool isReadOnly = row.Table.Columns["Id"].ReadOnly;
                    row.Table.Columns["Id"].ReadOnly = false;
                    row["Id"] = existingDict[key];
                    row.Table.Columns["Id"].ReadOnly = isReadOnly;
                }
            }
        }

        // ==========================================
        // UI 輔助方法
        // ==========================================
        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) {
                if (_dgv.Columns[e.ColumnIndex].Name == "鑑別日期") {
                    var cell = _dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (cell.Value == DBNull.Value || string.IsNullOrWhiteSpace(cell.Value?.ToString())) {
                        cell.Value = DateTime.Today.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        private DateTime GetStartDate() { return ParseComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddYears(-1)); }
        private DateTime GetEndDate() { return ParseComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today); }
        private DateTime ParseComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime defaultDate) {
            if (y.SelectedItem == null || m.SelectedItem == null || d.SelectedItem == null) return defaultDate;
            if (int.TryParse(y.SelectedItem.ToString(), out int year) &&
                int.TryParse(m.SelectedItem.ToString(), out int month) &&
                int.TryParse(d.SelectedItem.ToString(), out int day)) {
                try {
                    int daysInMonth = DateTime.DaysInMonth(year, month);
                    if (day > daysInMonth) day = daysInMonth;
                    return new DateTime(year, month, day);
                } catch { return defaultDate; }
            }
            return defaultDate;
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2");
            d.SelectedItem = date.Day.ToString("D2");
        }

        private void UpdateCboColumns()
        {
            string currentSelection = _cboSearchColumn.SelectedItem?.ToString();
            _cboColumns.Items.Clear();
            _cboSearchColumn.Items.Clear();
            _cboSearchColumn.Items.Add(""); 

            foreach (DataGridViewColumn c in _dgv.Columns) {
                if (c.Name != "Id") {
                    if (c.Name != "日期") _cboColumns.Items.Add(c.Name);
                    _cboSearchColumn.Items.Add(c.Name); 
                }
            }
            if (!string.IsNullOrEmpty(currentSelection) && _cboSearchColumn.Items.Contains(currentSelection)) {
                _cboSearchColumn.SelectedItem = currentSelection;
            } else {
                _cboSearchColumn.SelectedIndex = 0; 
            }
        }

        private void SetupComboBoxColumns()
        {
            ReplaceWithComboBox("類別", new string[] { "法律", "命令", "行政規則", "解釋令函", "" });
            ReplaceWithComboBox("適用性", new string[] { "適用", "不適用", "參考", "確認中", "" });

            string[] checkItems = new string[] { "", "v" };
            ReplaceWithComboBox("有提升績效機會", checkItems);
            ReplaceWithComboBox("有潛在不符合風險", checkItems);

            List<string> itemsTiao = new List<string> { "" };
            for (int i = 1; i <= 500; i++) itemsTiao.Add(i.ToString("D3"));
            ReplaceWithComboBox("條", itemsTiao.ToArray());

            List<string> itemsSmall = new List<string> { "" };
            for (int i = 1; i <= 20; i++) itemsSmall.Add(i.ToString("D2"));
            
            ReplaceWithComboBox("項", itemsSmall.ToArray());
            ReplaceWithComboBox("款", itemsSmall.ToArray());
            ReplaceWithComboBox("目", itemsSmall.ToArray());
        }

        private void ReplaceWithComboBox(string colName, string[] defaultItems)
        {
            if (_dgv.Columns.Contains(colName) && !(_dgv.Columns[colName] is DataGridViewComboBoxColumn))
            {
                int colIndex = _dgv.Columns[colName].Index;
                _dgv.Columns.Remove(colName);

                DataGridViewComboBoxColumn cboCol = new DataGridViewComboBoxColumn();
                cboCol.Name = colName;
                cboCol.HeaderText = colName;
                cboCol.DataPropertyName = colName; 
                
                List<string> finalItems = new List<string>(defaultItems);
                if (_dgv.DataSource is DataTable dt) {
                    foreach (DataRow row in dt.Rows) {
                        string val = row[colName]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(val) && !finalItems.Contains(val)) {
                            finalItems.Add(val);
                        }
                    }
                }
                cboCol.Items.AddRange(finalItems.ToArray());
                cboCol.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
                cboCol.FlatStyle = FlatStyle.Flat;
                cboCol.SortMode = DataGridViewColumnSortMode.Automatic; 

                _dgv.Columns.Insert(colIndex, cboCol);
            }
        }

        private void SetupTextWrapping()
        {
            Dictionary<string, int> columnWidths = new Dictionary<string, int> {
                { "法規名稱", 250 }, { "內容", 400 }, { "重點摘要", 200 }, { "備註", 150 }
            };
            foreach (var kvp in columnWidths) {
                if (_dgv.Columns.Contains(kvp.Key)) {
                    _dgv.Columns[kvp.Key].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    _dgv.Columns[kvp.Key].Width = kvp.Value; 
                    _dgv.Columns[kvp.Key].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }
        }

        private void Dgv_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox cbo) { cbo.DropDownStyle = ComboBoxStyle.DropDownList; }
            else if (e.Control is TextBox txt) { txt.Multiline = true; }
        }

        private void Dgv_DataError(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; }

        // ==========================================
        // 匯入與匯出邏輯 (包含防卡頓優化)
        // ==========================================
        private void BtnExport_Click(object sender, EventArgs e) {
            if (_dgv.Rows.Count <= 1) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = _tableName + "_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    using (ExcelPackage p = new ExcelPackage()) { var ws = p.Workbook.Worksheets.Add("Data"); ws.Cells["A1"].LoadFromDataTable((DataTable)_dgv.DataSource, true); ws.Cells.AutoFitColumns(); p.SaveAs(new FileInfo(sfd.FileName)); }
                    MessageBox.Show("資料匯出成功！");
                }
            }
        }

        private async void BtnImportExcel_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "請選擇要匯入的 Excel 檔案" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                        SetUIState(false, "Excel 解析中，請稍候...", Color.Orange);

                        DataTable originalDt = (DataTable)_dgv.DataSource;
                        _dgv.DataSource = null; 
                        DataTable tempDt = originalDt.Clone();

                        await Task.Run(() => {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                if (ws == null || ws.Dimension == null) return;

                                int rowCount = ws.Dimension.Rows;
                                int colCount = ws.Dimension.Columns;

                                string[] headers = new string[colCount];
                                for (int c = 1; c <= colCount; c++) headers[c - 1] = ws.Cells[1, c].Text.Trim();

                                for (int r = 2; r <= rowCount; r++) {
                                    DataRow nr = tempDt.NewRow();
                                    bool hasData = false;
                                    for (int c = 1; c <= colCount; c++) {
                                        string cn = headers[c - 1];
                                        if (tempDt.Columns.Contains(cn) && cn != "Id") {
                                            string val = ws.Cells[r, c].Text.Trim();
                                            nr[cn] = val;
                                            if (!string.IsNullOrEmpty(val)) hasData = true;
                                        }
                                    }
                                    if (hasData) tempDt.Rows.Add(nr);
                                }
                                EnforceDateFormats(tempDt); // 🟢 匯入後強制統一日期
                            }
                        });

                        originalDt.Merge(tempDt);
                        _dgv.DataSource = originalDt; 
                        ApplyGridFormatting();

                        SetUIState(true, $"匯入成功！新增資料後總筆數：{tempDt.Rows.Count}", Color.Green);
                        MessageBox.Show($"載入 {tempDt.Rows.Count} 筆資料成功！\n系統已就緒，請點擊「儲存」以寫入資料庫。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        await LoadGridDataAsync(false); 
                        MessageBox.Show("匯入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } finally {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void BtnRtfToExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "RTF 法規檔案 (*.rtf)|*.rtf", Title = "請選擇全國法規資料庫下載的 RTF 檔案" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + "_轉換.xlsx" }) {
                        if (sfd.ShowDialog() == DialogResult.OK) {
                            try {
                                LawRtfToExcelConverter.Convert(ofd.FileName, sfd.FileName);
                                MessageBox.Show("轉換成功！\n您現在可以點擊「匯入 EXCEL」將產生的檔案載入系統。", "轉換完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            } catch (Exception ex) {
                                MessageBox.Show("轉換失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.S) {
                e.Handled = true; e.SuppressKeyPress = true; _btnSave?.PerformClick(); 
            }
            else if (e.Control && e.KeyCode == Keys.V) {
                try {
                    string text = Clipboard.GetText(); if (string.IsNullOrEmpty(text)) return;
                    _dgv.SuspendLayout(); 

                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    int r = _dgv.CurrentCell.RowIndex, c = _dgv.CurrentCell.ColumnIndex;
                    DataTable dt = (DataTable)_dgv.DataSource;

                    foreach (string line in lines) {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++) {
                            if (c + i < _dgv.Columns.Count && !_dgv.Columns[c + i].ReadOnly) {
                                _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                            }
                        }
                        r++;
                    }
                    EnforceDateFormats(dt); // 🟢 貼上後立即修正日期防呆
                    _dgv.ResumeLayout();
                } catch (Exception ex) { 
                    _dgv.ResumeLayout(); MessageBox.Show("貼上失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
