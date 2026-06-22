/// FILE: Safety_System/Dashboard/App_AuditDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml; 
using System.IO;
using System.Threading.Tasks;

namespace Safety_System
{
    public class App_AuditDashboard
    {
        // UI 控制項 - 查詢區
        private DateTimePicker _dtpStartDate;
        private NumericUpDown _numDays;
        private ComboBox _cboDb;
        private ComboBox _cboTable;
        private ComboBox _cboColumn;
        private TextBox _txtFormNumbers;
        private Button _btnSearch;
        private Button _btnExportExcel;

        // UI 控制項 - 結果區 A (日期區間)
        private Label _lblResultATitle;
        private DataGridView _dgvResultA;
        private Button _btnSettingsA;
        private Button _btnSaveA;

        // UI 控制項 - 結果區 B (單號追蹤)
        private Label _lblResultBTitle;
        private DataGridView _dgvResultB;
        private Button _btnSettingsB;
        private Button _btnSaveB;

        // 資料庫快取
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 預設顯示欄位設定
        private readonly string[] _defaultVisibleCols = { "日期", "單位", "表單單號", "表單主題", "建議改善事項", "追蹤改善狀況", "改善進度", "缺失責任人", "附件檔案", "備註", "連結" };
        
        // 允許編輯的欄位名單
        private readonly string[] _editableCols = { "單位", "缺失責任人", "建議改善事項", "追蹤改善狀況", "改善進度", "附件檔案", "備註" };

        private ContextMenuStrip _ctxMenuA;
        private ContextMenuStrip _ctxMenuB;
        private int _rightClickedColIndexA = -1;
        private int _rightClickedColIndexB = -1;

        // 寬度記憶快取
        private Dictionary<string, int> _columnWidthsA = new Dictionary<string, int>();
        private Dictionary<string, int> _columnWidthsB = new Dictionary<string, int>();

        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        public Control GetView()
        {
            _dbMap = App_DbConfig.GetDbMapCache();

            Panel mainScrollPanel = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.WhiteSmoke, 
                AutoScroll = true, 
                Padding = new Padding(20) 
            };

            TableLayoutPanel masterLayout = new TableLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                ColumnCount = 1, 
                RowCount = 4,
                Margin = new Padding(0)
            };
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            masterLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            // ==========================================
            // 第一行：大標題
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, Height = 60, Margin = new Padding(0) };
            Label lblTitle = new Label { Text = "🛡️ 稽核資料查詢與追蹤", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            pnlHeader.Controls.Add(lblTitle);
            masterLayout.Controls.Add(pnlHeader, 0, 0);

            // ==========================================
            // 第一大框：查詢設定
            // ==========================================
            GroupBox boxQuery = new GroupBox { Text = "⚙️ 查詢設定", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0, 0, 0, 20) };
            
            FlowLayoutPanel flpQueryMain = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };

            // A. 日期區間設定
            FlowLayoutPanel flpDate = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 5, 0, 15) };
            _dtpStartDate = new DateTimePicker { Format = DateTimePickerFormat.Short, Font = new Font("Microsoft JhengHei UI", 12F), Width = 150, Margin = new Padding(5, 4, 20, 0), Value = DateTime.Today.AddDays(-7) };
            _numDays = new NumericUpDown { Minimum = 0, Maximum = 1000, Value = 7, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80, Margin = new Padding(5, 4, 5, 0) };
            
            flpDate.Controls.AddRange(new Control[] {
                new Label { Text = "【A. 查詢區間設定】", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.Teal, Margin = new Padding(0, 6, 10, 0) },
                new Label { Text = "查詢日期起:", AutoSize = true, Margin = new Padding(0, 6, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _dtpStartDate,
                new Label { Text = "查詢天數:", AutoSize = true, Margin = new Padding(0, 6, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _numDays,
                new Label { Text = "天", AutoSize = true, Margin = new Padding(0, 6, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) }
            });

            // B. 資料來源與單號設定
            FlowLayoutPanel flpSource = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 5, 0, 10) };
            _cboDb = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 160, Margin = new Padding(5, 4, 15, 0) };
            _cboTable = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 220, Margin = new Padding(5, 4, 15, 0) };
            _cboColumn = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 200, Margin = new Padding(5, 4, 20, 0) };

            flpSource.Controls.AddRange(new Control[] {
                new Label { Text = "【B. 單號追蹤設定】", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.Chocolate, Margin = new Padding(0, 6, 10, 0) },
                new Label { Text = "資料庫:", AutoSize = true, Margin = new Padding(0, 6, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _cboDb,
                new Label { Text = "資料表:", AutoSize = true, Margin = new Padding(0, 6, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _cboTable,
                new Label { Text = "資料欄:", AutoSize = true, Margin = new Padding(0, 6, 0, 0), Font = new Font("Microsoft JhengHei UI", 12F) }, _cboColumn
            });

            // 多行文字框與按鈕
            FlowLayoutPanel flpText = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 5, 0, 5) };
            _txtFormNumbers = new TextBox { Multiline = true, Width = 350, Height = 100, Font = new Font("Consolas", 12F), ScrollBars = ScrollBars.Vertical, Margin = new Padding(5, 0, 20, 0) };
            
            _btnSearch = new Button { Text = "🚀 執行雙向查詢", Size = new Size(200, 60), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 20, 0, 0) };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += async (s, e) => await ExecuteSearchAsync();

            _btnExportExcel = new Button { Text = "📤 匯出 Excel", Size = new Size(180, 60), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(15, 20, 0, 0) };
            _btnExportExcel.FlatAppearance.BorderSize = 0;
            _btnExportExcel.Click += (s, e) => ExportToExcel();

            flpText.Controls.AddRange(new Control[] {
                new Label { Text = "查詢巡檢表單單號:\n(每行代表一段查詢文字)", AutoSize = true, Margin = new Padding(30, 5, 0, 0), Font = new Font("Microsoft JhengHei UI", 11F), ForeColor = Color.DimGray },
                _txtFormNumbers,
                _btnSearch,
                _btnExportExcel
            });

            flpQueryMain.Controls.Add(flpDate);
            flpQueryMain.Controls.Add(flpSource);
            flpQueryMain.Controls.Add(flpText);
            boxQuery.Controls.Add(flpQueryMain);
            masterLayout.Controls.Add(boxQuery, 0, 1);

            // ==========================================
            // 第三個框：取得查詢資料 (A. 日期區間)
            // ==========================================
            GroupBox boxResultA = new GroupBox { Text = "📊 查詢資料 (依日期區間)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0, 0, 0, 20) };
            
            Panel pnlHeaderA = new Panel { Dock = DockStyle.Top, Height = 45 };
            _lblResultATitle = new Label { Text = "查詢區間：尚未查詢", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Left, AutoSize = true, Padding = new Padding(0,5,0,0) };
            
            _btnSaveA = new Button { Text = "💾 儲存變更", Size = new Size(130, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Right, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0,0,10,0) };
            _btnSaveA.FlatAppearance.BorderSize = 0;
            _btnSaveA.Click += (s, e) => SaveGridData(_dgvResultA);

            _btnSettingsA = new Button { Text = "⚙️ 顯示與排序設定", Size = new Size(180, 35), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Right, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSettingsA.FlatAppearance.BorderSize = 0;
            _btnSettingsA.Click += (s, e) => OpenGridSettings(_dgvResultA, "GridA");

            pnlHeaderA.Controls.Add(_lblResultATitle);
            pnlHeaderA.Controls.Add(_btnSaveA);
            pnlHeaderA.Controls.Add(new Panel { Width = 15, Dock=DockStyle.Right}); 
            pnlHeaderA.Controls.Add(_btnSettingsA);

            _dgvResultA = CreateStandardGrid("GridA");
            
            boxResultA.Controls.Add(_dgvResultA);
            boxResultA.Controls.Add(pnlHeaderA);
            masterLayout.Controls.Add(boxResultA, 0, 2);

            // ==========================================
            // 第四個框：取得查詢資料 (B. 單號追蹤)
            // ==========================================
            GroupBox boxResultB = new GroupBox { Text = "📑 查詢資料 (依單號追蹤)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0, 0, 0, 20) };
            
            Panel pnlHeaderB = new Panel { Dock = DockStyle.Top, Height = 45 };
            _lblResultBTitle = new Label { Text = "巡檢異常改善單 - 追蹤", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.Chocolate, Dock = DockStyle.Left, AutoSize = true, Padding = new Padding(0,5,0,0)  };
            
            _btnSaveB = new Button { Text = "💾 儲存變更", Size = new Size(130, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Right, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0,0,10,0) };
            _btnSaveB.FlatAppearance.BorderSize = 0;
            _btnSaveB.Click += (s, e) => SaveGridData(_dgvResultB);

            _btnSettingsB = new Button { Text = "⚙️ 顯示與排序設定", Size = new Size(180, 35), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Right, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSettingsB.FlatAppearance.BorderSize = 0;
            _btnSettingsB.Click += (s, e) => OpenGridSettings(_dgvResultB, "GridB");

            pnlHeaderB.Controls.Add(_lblResultBTitle);
            pnlHeaderB.Controls.Add(_btnSaveB);
            pnlHeaderB.Controls.Add(new Panel { Width = 15, Dock=DockStyle.Right}); 
            pnlHeaderB.Controls.Add(_btnSettingsB);

            _dgvResultB = CreateStandardGrid("GridB");
            
            boxResultB.Controls.Add(_dgvResultB);
            boxResultB.Controls.Add(pnlHeaderB);
            masterLayout.Controls.Add(boxResultB, 0, 3);

            mainScrollPanel.Controls.Add(masterLayout);

            InitDropdownLogic();
            InitContextMenus();

            return mainScrollPanel;
        }

        private DataGridView CreateStandardGrid(string gridId)
        {
            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Top, 
                Height = 350,
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                AllowUserToDeleteRows = false, 
                ReadOnly = false, // 🟢 開放編輯
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, // 🟢 讓使用者可調寬度
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                Font = new Font("Microsoft JhengHei UI", 11F), 
                BorderStyle = BorderStyle.FixedSingle, 
                Margin = new Padding(0, 10, 0, 0),
                AllowUserToOrderColumns = true, // 🟢 允許拖曳排序
                AllowUserToResizeColumns = true // 🟢 允許調整寬度
            };
            
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.EditMode = DataGridViewEditMode.EditOnEnter;

            typeof(DataGridView).InvokeMember("DoubleBuffered", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, 
                null, dgv, new object[] { true });

            // 綁定事件
            dgv.CellClick += (s, e) => Dgv_CellClick(s, e, dgv);
            dgv.CellMouseClick += (s, e) => {
                if (e.Button == MouseButtons.Right && e.ColumnIndex >= 0 && e.RowIndex >= 0) {
                    if (gridId == "GridA") _rightClickedColIndexA = e.ColumnIndex;
                    else _rightClickedColIndexB = e.ColumnIndex;
                    
                    dgv.ClearSelection();
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                    dgv.CurrentCell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    
                    if (gridId == "GridA") _ctxMenuA.Show(Cursor.Position);
                    else _ctxMenuB.Show(Cursor.Position);
                }
            };

            // 記憶排序與寬度
            dgv.ColumnWidthChanged += (s, e) => {
                if (e.Column != null && e.Column.Visible && e.Column.Width > 0) {
                    var dict = gridId == "GridA" ? _columnWidthsA : _columnWidthsB;
                    dict[e.Column.Name] = e.Column.Width;
                    DataManager.SaveGridConfig("AuditDash", gridId, "Width", e.Column.Name, e.Column.Width.ToString());
                }
            };

            dgv.ColumnDisplayIndexChanged += (s, e) => {
                try {
                    List<string> orderedCols = new List<string>();
                    foreach (DataGridViewColumn c in dgv.Columns) orderedCols.Add(c.Name);
                    orderedCols.Sort((a, b) => dgv.Columns[a].DisplayIndex.CompareTo(dgv.Columns[b].DisplayIndex));
                    DataManager.SaveGridConfig("AuditDash", gridId, "Order", "All", string.Join(",", orderedCols)); 
                } catch { }
            };

            return dgv;
        }

        private void InitContextMenus()
        {
            _ctxMenuA = new ContextMenuStrip { Font = new Font("Microsoft JhengHei UI", 11F) };
            var itemUrlA = new ToolStripMenuItem("🌐 開啟網址");
            itemUrlA.Click += (s, e) => OpenUrl(_dgvResultA, _rightClickedColIndexA);
            _ctxMenuA.Items.Add(itemUrlA);

            _ctxMenuB = new ContextMenuStrip { Font = new Font("Microsoft JhengHei UI", 11F) };
            var itemUrlB = new ToolStripMenuItem("🌐 開啟網址");
            itemUrlB.Click += (s, e) => OpenUrl(_dgvResultB, _rightClickedColIndexB);
            _ctxMenuB.Items.Add(itemUrlB);
        }

        private void OpenUrl(DataGridView dgv, int colIndex)
        {
            if (colIndex >= 0 && dgv.CurrentCell != null)
            {
                int r = dgv.CurrentCell.RowIndex;
                object val = dgv.Rows[r].Cells[colIndex].Value;
                if (val != null)
                {
                    string url = val.ToString().Trim();
                    if (string.IsNullOrWhiteSpace(url)) return;

                    if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && 
                        !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    {
                        url = "http://" + url;
                    }

                    try { System.Diagnostics.Process.Start(url); } 
                    catch (Exception ex) {
                        MessageBox.Show($"無法開啟網址！\n請確認網址格式是否正確：\n\n{url}\n\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void InitDropdownLogic()
        {
            _cboDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            foreach (var kvp in _dbMap) {
                _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
            }

            _cboDb.SelectedIndexChanged += (s, e) => {
                _cboTable.Items.Clear();
                _cboColumn.Items.Clear();
                _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
                
                var selDb = _cboDb.SelectedItem as ItemMap;
                if (selDb != null && !string.IsNullOrEmpty(selDb.EnName) && _dbMap.ContainsKey(selDb.EnName)) {
                    foreach (var tb in _dbMap[selDb.EnName].Tables) {
                        _cboTable.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                    }
                }
                if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
            };

            _cboTable.SelectedIndexChanged += (s, e) => {
                _cboColumn.Items.Clear();
                var selDb = _cboDb.SelectedItem as ItemMap;
                var selTb = _cboTable.SelectedItem as ItemMap;

                if (selDb != null && selTb != null && !string.IsNullOrEmpty(selDb.EnName) && !string.IsNullOrEmpty(selTb.EnName)) {
                    var cols = DataManager.GetColumnNames(selDb.EnName, selTb.EnName);
                    foreach (var c in cols) {
                        if (c != "Id" && c != "附件檔案" && c != "備註") _cboColumn.Items.Add(c);
                    }
                }
                
                bool preSelected = false;
                for (int i = 0; i < _cboColumn.Items.Count; i++) {
                    string colText = _cboColumn.Items[i].ToString();
                    if (colText.Contains("單號") || colText.Contains("編號")) {
                        _cboColumn.SelectedIndex = i;
                        preSelected = true;
                        break;
                    }
                }
                if (!preSelected && _cboColumn.Items.Count > 0) _cboColumn.SelectedIndex = 0;
            };

            for (int i = 0; i < _cboDb.Items.Count; i++) {
                if (((ItemMap)_cboDb.Items[i]).EnName == "Safety") {
                    _cboDb.SelectedIndex = i; break;
                }
            }
            if (_cboDb.SelectedIndex > 0) {
                for (int i = 0; i < _cboTable.Items.Count; i++) {
                    if (((ItemMap)_cboTable.Items[i]).EnName == "SafetyInspection") {
                        _cboTable.SelectedIndex = i; break;
                    }
                }
            }
        }

        private async Task ExecuteSearchAsync()
        {
            var selDb = _cboDb.SelectedItem as ItemMap;
            var selTb = _cboTable.SelectedItem as ItemMap;
            
            if (selDb == null || selTb == null || string.IsNullOrEmpty(selDb.EnName) || string.IsNullOrEmpty(selTb.EnName)) {
                MessageBox.Show("請先選擇資料庫與資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string dbName = selDb.EnName;
            string tbName = selTb.EnName;
            string tbChName = selTb.ChName;
            string colName = _cboColumn.SelectedItem?.ToString() ?? "";

            DateTime sDate = _dtpStartDate.Value.Date;
            DateTime eDate = sDate.AddDays((double)_numDays.Value);

            _lblResultATitle.Text = $"查詢區間：{sDate:yyyy年MM月dd日} ~ {eDate:yyyy年MM月dd日} ~ {tbChName}";

            _btnSearch.Enabled = false;
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

            DataTable rawData = null;
            DataView dvA = null;
            DataView dvB = null;

            await Task.Run(() => {
                rawData = DataManager.GetTableData(dbName, tbName, "", "", "");
                if (rawData == null || rawData.Rows.Count == 0) return;

                dvA = new DataView(rawData);
                string dateCol = GetDateColumn(rawData);
                if (!string.IsNullOrEmpty(dateCol)) {
                    dvA.RowFilter = $"[{dateCol}] >= '{sDate:yyyy-MM-dd}' AND [{dateCol}] <= '{eDate:yyyy-MM-dd}'";
                }

                dvB = new DataView(rawData);
                string[] forms = _txtFormNumbers.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (forms.Length > 0 && !string.IsNullOrEmpty(colName) && rawData.Columns.Contains(colName)) {
                    string inClause = string.Join(",", forms.Select(f => $"'{f.Trim().Replace("'", "''")}'"));
                    dvB.RowFilter = $"[{colName}] IN ({inClause})";
                } else {
                    dvB.RowFilter = "1=0"; 
                }
            });

            if (rawData != null) {
                // ToTable() 建立的 DataTable 會切斷與原底層的連結，所以呼叫 AcceptChanges 讓它們變成 Unchanged
                DataTable dtA = dvA.ToTable(); dtA.AcceptChanges();
                DataTable dtB = dvB.ToTable(); dtB.AcceptChanges();
                
                _dgvResultA.DataSource = dtA;
                _dgvResultB.DataSource = dtB;

                ApplyGridSettings(_dgvResultA, "GridA");
                ApplyGridSettings(_dgvResultB, "GridB");
            } else {
                _dgvResultA.DataSource = null;
                _dgvResultB.DataSource = null;
            }

            _btnSearch.Enabled = true;
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private string GetDateColumn(DataTable dt)
        {
            if (dt.Columns.Contains("日期")) return "日期";
            if (dt.Columns.Contains("年月")) return "年月";
            if (dt.Columns.Contains("清運日期")) return "清運日期";
            return "";
        }

        // ==========================================
        // 欄位編輯與下拉選單
        // ==========================================
        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e, DataGridView dgv)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && !dgv.Rows[e.RowIndex].IsNewRow) 
            {
                string colName = dgv.Columns[e.ColumnIndex].Name;
                if (!_editableCols.Contains(colName)) return; // 非編輯欄位不理會

                if (colName == "附件檔案") 
                {
                    string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
                    string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;

                    string currentVal = dgv[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";
                    
                    // 嘗試抓日期作資料夾分類
                    string dateCol = GetDateColumn((DataTable)dgv.DataSource);
                    string rowDateStr = "";
                    if (!string.IsNullOrEmpty(dateCol) && dgv.Columns.Contains(dateCol)) {
                        rowDateStr = dgv[dateCol, e.RowIndex].Value?.ToString() ?? "";
                    }
                    string targetFolder = string.IsNullOrEmpty(rowDateStr) ? DateTime.Now.ToString("yyyy-MM") : (rowDateStr.Length >= 7 ? rowDateStr.Substring(0, 7) : rowDateStr);

                    using (var frm = new AttachmentManagerUI(currentVal, dbName, tbName, targetFolder, delegate(string path) { DeletePhysicalFile(path, dgv, e.RowIndex); })) {
                        if (frm.ShowDialog() == DialogResult.OK) { 
                            dgv[e.ColumnIndex, e.RowIndex].Value = frm.FinalPathsString; 
                            dgv.EndEdit(); 
                        }
                    }
                } 
                else 
                {
                    string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;
                    string prefix = $"{tbName}|{colName}|";
                    bool isSingleDropdown = App_DropdownManager.DropdownCache.Keys.Any(k => k.StartsWith(prefix));
                    bool isMultiDropdown = App_DropdownManager.MultiSelectCache.ContainsKey($"{tbName}|{colName}");

                    if (isSingleDropdown) {
                        ShowCustomDropdown(dgv, tbName, e.RowIndex, e.ColumnIndex);
                    } else if (isMultiDropdown) {
                        TriggerMultiSelectDialog(dgv, tbName, e.RowIndex, e.ColumnIndex);
                    }
                }
            }
        }

        private void TriggerMultiSelectDialog(DataGridView dgv, string tbName, int rowIndex, int colIndex)
        {
            string colName = dgv.Columns[colIndex].Name;
            string multiKey = $"{tbName}|{colName}";

            if (App_DropdownManager.MultiSelectCache.ContainsKey(multiKey))
            {
                var options = App_DropdownManager.MultiSelectCache[multiKey];
                string currentVal = dgv[colIndex, rowIndex].Value?.ToString() ?? "";

                using (Form fMulti = new Form())
                {
                    fMulti.Text = $"複選組合：{colName}";
                    fMulti.Size = new Size(380, 500); 
                    fMulti.StartPosition = FormStartPosition.CenterParent;
                    fMulti.FormBorderStyle = FormBorderStyle.FixedDialog;
                    fMulti.MaximizeBox = false;
                    fMulti.MinimizeBox = false;
                    fMulti.BackColor = Color.White;

                    DataGridView dgvMulti = new DataGridView {
                        Dock = DockStyle.Fill,
                        BackgroundColor = Color.White,
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        AllowUserToResizeColumns = false,
                        AllowUserToResizeRows = false,
                        RowHeadersVisible = false,
                        ColumnHeadersVisible = false,
                        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                        CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                        Font = new Font("Microsoft JhengHei UI", 12F)
                    };
                    dgvMulti.RowTemplate.Height = 40; 

                    DataGridViewCheckBoxColumn chkCol = new DataGridViewCheckBoxColumn { Name = "Check", Width = 40, AutoSizeMode = DataGridViewAutoSizeColumnMode.None };
                    DataGridViewImageColumn imgCol = new DataGridViewImageColumn { Name = "Icon", Width = 40, AutoSizeMode = DataGridViewAutoSizeColumnMode.None, ImageLayout = DataGridViewImageCellLayout.Zoom };
                    DataGridViewTextBoxColumn txtCol = new DataGridViewTextBoxColumn { Name = "Text", ReadOnly = true };

                    dgvMulti.Columns.Add(chkCol);
                    dgvMulti.Columns.Add(imgCol);
                    dgvMulti.Columns.Add(txtCol);

                    string[] currentSelected = currentVal.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
                    
                    foreach (var opt in options)
                    {
                        int rIdx = dgvMulti.Rows.Add();
                        bool isChecked = currentSelected.Contains(opt.Text);
                        dgvMulti.Rows[rIdx].Cells["Check"].Value = isChecked;
                        dgvMulti.Rows[rIdx].Cells["Icon"].Value = opt.GetImage() ?? new Bitmap(1, 1); 
                        dgvMulti.Rows[rIdx].Cells["Text"].Value = opt.Text;
                        dgvMulti.Rows[rIdx].Tag = opt.Text; 
                    }

                    dgvMulti.CellClick += (s, ev) => {
                        if (ev.RowIndex >= 0) {
                            bool current = Convert.ToBoolean(dgvMulti.Rows[ev.RowIndex].Cells["Check"].Value);
                            dgvMulti.Rows[ev.RowIndex].Cells["Check"].Value = !current;
                        }
                    };

                    Button btnOk = new Button {
                        Text = "✔️ 確認並填入",
                        Dock = DockStyle.Bottom,
                        Height = 45,
                        BackColor = Color.SteelBlue,
                        ForeColor = Color.White,
                        Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                        Cursor = Cursors.Hand
                    };
                    
                    btnOk.Click += (s2, ev2) => {
                        List<string> selectedItems = new List<string>();
                        foreach (DataGridViewRow r in dgvMulti.Rows) {
                            if (Convert.ToBoolean(r.Cells["Check"].Value)) {
                                selectedItems.Add(r.Tag.ToString());
                            }
                        }
                        
                        dgv[colIndex, rowIndex].Value = string.Join(Environment.NewLine, selectedItems);
                        dgv.EndEdit();
                        fMulti.DialogResult = DialogResult.OK;
                    };

                    fMulti.Controls.Add(dgvMulti);
                    fMulti.Controls.Add(btnOk);
                    fMulti.ShowDialog(Form.ActiveForm);
                }
            }
        }

        private void ShowCustomDropdown(DataGridView dgv, string tbName, int rowIndex, int colIndex)
        {
            string colName = dgv.Columns[colIndex].Name;
            string parentColName = "";
            
            foreach (var kvp in App_DropdownManager.DropdownCache) {
                string[] parts = kvp.Key.Split('|');
                if (parts.Length == 4 && parts[0] == tbName && parts[1] == colName && !string.IsNullOrEmpty(parts[2])) {
                    parentColName = parts[2]; break;
                }
            }

            List<DropdownItemDef> items = new List<DropdownItemDef>();

            if (!string.IsNullOrEmpty(parentColName) && dgv.Columns.Contains(parentColName)) {
                string parentVal = dgv.Rows[rowIndex].Cells[parentColName].Value?.ToString() ?? "";
                string key = $"{tbName}|{colName}|{parentColName}|{parentVal}";
                if (App_DropdownManager.DropdownCache.ContainsKey(key)) {
                    items = App_DropdownManager.DropdownCache[key];
                }
            } else {
                string prefix = $"{tbName}|{colName}|";
                foreach(var kvp in App_DropdownManager.DropdownCache) {
                    if (kvp.Key.StartsWith(prefix)) {
                        items.AddRange(kvp.Value);
                    }
                }
            }

            var uniqueItems = items.GroupBy(x => x.Text).Select(g => g.First()).ToList();
            if (uniqueItems.Count == 0) return;

            ToolStripDropDown dropDown = new ToolStripDropDown();
            dropDown.Margin = Padding.Empty;
            dropDown.Padding = Padding.Empty;

            ListBox listBox = new ListBox {
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 36, 
                Font = new Font("Microsoft JhengHei UI", 12F),
                BorderStyle = BorderStyle.FixedSingle
            };

            Rectangle cellRect = dgv.GetCellDisplayRectangle(colIndex, rowIndex, false);
            
            int maxTextWidth = 180;
            using (Graphics g = dgv.CreateGraphics()) {
                foreach (var item in uniqueItems) {
                    int w = (int)g.MeasureString(item.Text, listBox.Font).Width;
                    if (w + 60 > maxTextWidth) maxTextWidth = w + 60; 
                }
            }

            int finalWidth = Math.Max(cellRect.Width, maxTextWidth);
            int finalHeight = uniqueItems.Count * 36 + 5;
            if (finalHeight > 300) finalHeight = 300; 

            listBox.Width = finalWidth;
            listBox.Height = finalHeight;

            listBox.DrawItem += (s, e) => {
                if (e.Index < 0) return;
                e.DrawBackground();

                var item = uniqueItems[e.Index];
                Brush textBrush = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? Brushes.White : Brushes.Black;
                
                int imgSize = 24; 
                int textOffset = 8; 

                Image img = item.GetImage();
                if (img != null) {
                    int imgY = e.Bounds.Y + (e.Bounds.Height - imgSize) / 2;
                    e.Graphics.DrawImage(img, e.Bounds.X + 8, imgY, imgSize, imgSize);
                    textOffset = 40; 
                }

                e.Graphics.DrawString(item.Text, listBox.Font, textBrush, new PointF(e.Bounds.X + textOffset, e.Bounds.Y + 8));
                e.DrawFocusRectangle();
            };

            foreach (var item in uniqueItems) listBox.Items.Add(item);

            listBox.MouseClick += (s, e) => {
                if (listBox.SelectedIndex >= 0) {
                    dgv[colIndex, rowIndex].Value = uniqueItems[listBox.SelectedIndex].Text;
                    dgv.EndEdit();
                    dropDown.Close();
                }
            };

            ToolStripControlHost host = new ToolStripControlHost(listBox);
            host.Margin = Padding.Empty;
            host.Padding = Padding.Empty;
            host.AutoSize = false;
            host.Size = new Size(finalWidth, finalHeight);
            
            dropDown.Items.Add(host);

            Point displayLocation = dgv.PointToScreen(new Point(cellRect.Left, cellRect.Bottom));
            dropDown.Show(displayLocation);
            
            listBox.Focus();
        }

        private void DeletePhysicalFile(string relativePath, DataGridView dgv, int currentRowIndex) 
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return;
            bool isUsedByOthers = false;
            
            foreach (DataGridViewRow row in dgv.Rows) {
                if (row.Index == currentRowIndex || row.IsNewRow) continue;
                if (dgv.Columns.Contains("附件檔案")) {
                    object cellVal = row.Cells["附件檔案"].Value;
                    if (cellVal != null) { 
                        string[] paths = cellVal.ToString().Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries); 
                        foreach(string p in paths) { if (p == relativePath) { isUsedByOthers = true; break; } }
                    }
                }
            }
            
            if (isUsedByOthers) return;
            
            try {
                string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                if (File.Exists(absPath)) File.Delete(absPath); 
            } catch { }
        }

        private void SaveGridData(DataGridView dgv) 
        {
            dgv.EndEdit();
            if (dgv.DataSource is DataTable dt) {
                DataTable changes = dt.GetChanges();
                if (changes != null && changes.Rows.Count > 0) {
                    string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
                    string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;
                    
                    foreach(DataRow row in changes.Rows) {
                        DataManager.UpsertRecord(dbName, tbName, row);
                    }
                    dt.AcceptChanges();
                    MessageBox.Show("儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("沒有需要儲存的變更。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // ==========================================
        // 欄位顯示與排序設定系統 (結合拖曳與上下按鍵)
        // ==========================================
        private void ApplyGridSettings(DataGridView dgv, string gridId)
        {
            if (dgv.Columns.Count == 0) return;

            var visibilityDict = DataManager.LoadGridConfig("AuditDash", gridId, "Visibility");
            var orderDict = DataManager.LoadGridConfig("AuditDash", gridId, "Order");
            var widthDict = DataManager.LoadGridConfig("AuditDash", gridId, "Width");
            
            bool hasVisConfig = visibilityDict.Count > 0;
            bool hasOrdConfig = orderDict.ContainsKey("All");

            foreach (DataGridViewColumn col in dgv.Columns) {
                if (col.Name == "Id") { col.Visible = false; continue; }

                if (hasVisConfig && visibilityDict.ContainsKey(col.Name)) {
                    col.Visible = (visibilityDict[col.Name] == "1");
                } else {
                    bool isVisible = _defaultVisibleCols.Contains(col.Name);
                    
                    // A區(GridA)的資料顯示，預設把「缺失責任人」、「改善進度」隱藏
                    if (gridId == "GridA" && (col.Name == "缺失責任人" || col.Name == "改善進度")) {
                        isVisible = false;
                    }
                    // B區(GridB)的資料顯示，預設把「建議改善事項」隱藏
                    else if (gridId == "GridB" && col.Name == "建議改善事項") {
                        isVisible = false;
                    }
                    
                    col.Visible = isVisible;
                }

                // 開放編輯特定欄位
                col.ReadOnly = !_editableCols.Contains(col.Name);
                if (!col.ReadOnly && col.Name != "附件檔案") {
                    col.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 230); // 淺黃色代表可編輯
                }
                if (col.Name == "附件檔案") {
                    col.DefaultCellStyle.ForeColor = Color.Blue;
                    col.DefaultCellStyle.Font = new Font(dgv.Font, FontStyle.Underline);
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }

                // 恢復寬度記憶
                if (widthDict.ContainsKey(col.Name) && int.TryParse(widthDict[col.Name], out int w)) {
                    col.Width = w;
                }
            }

            if (hasOrdConfig) {
                string[] savedOrder = orderDict["All"].Split(',');
                for (int i = 0; i < savedOrder.Length; i++) {
                    if (dgv.Columns.Contains(savedOrder[i])) {
                        dgv.Columns[savedOrder[i]].DisplayIndex = i;
                    }
                }
            }
        }

        private void OpenGridSettings(DataGridView targetDgv, string gridId)
        {
            if (targetDgv.Columns.Count == 0) {
                MessageBox.Show("請先執行查詢產生資料後，再進行欄位設定。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (Form f = new Form { Text = "👁️ 顯示欄位與排序設定", Size = new Size(500, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White }) 
            {
                Label lblTop = new Label { 
                    Text = "勾選顯示項目，並可透過拖曳或右側按鈕調整排列順序：", 
                    Dock = DockStyle.Top, Padding = new Padding(10, 15, 10, 15), 
                    Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.SteelBlue,
                    AutoSize = true 
                }; 
                f.Controls.Add(lblTop);

                DataGridView dgvSettings = new DataGridView {
                    Dock = DockStyle.Fill,
                    BackgroundColor = Color.WhiteSmoke,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AllowUserToResizeColumns = false,
                    AllowUserToResizeRows = false,
                    RowHeadersVisible = false,
                    ColumnHeadersVisible = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                    MultiSelect = false,
                    AllowDrop = true,
                    Font = new Font("Microsoft JhengHei UI", 12F)
                };
                dgvSettings.RowTemplate.Height = 40;

                DataGridViewCheckBoxColumn colChk = new DataGridViewCheckBoxColumn { Name = "Visible", Width = 50 };
                DataGridViewTextBoxColumn colNameCol = new DataGridViewTextBoxColumn { Name = "ColName", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill };
                dgvSettings.Columns.Add(colChk);
                dgvSettings.Columns.Add(colNameCol);

                var sortedCols = targetDgv.Columns.Cast<DataGridViewColumn>()
                                          .Where(c => c.Name != "Id")
                                          .OrderBy(c => c.DisplayIndex)
                                          .ToList();

                foreach (var col in sortedCols) {
                    dgvSettings.Rows.Add(col.Visible, col.Name);
                }

                int dragFromIdx = -1;
                int dragToIdx = -1;
                Rectangle dragBox = Rectangle.Empty;

                dgvSettings.MouseDown += (s, e) => {
                    var hit = dgvSettings.HitTest(e.X, e.Y);
                    if (hit.RowIndex >= 0) {
                        dragFromIdx = hit.RowIndex;
                        Size dragSize = SystemInformation.DragSize;
                        dragBox = new Rectangle(new Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize);
                    } else {
                        dragBox = Rectangle.Empty;
                    }
                };

                dgvSettings.MouseMove += (s, e) => {
                    if ((e.Button & MouseButtons.Left) == MouseButtons.Left) {
                        if (dragBox != Rectangle.Empty && !dragBox.Contains(e.X, e.Y)) {
                            dgvSettings.DoDragDrop(dgvSettings.Rows[dragFromIdx], DragDropEffects.Move);
                        }
                    }
                };

                dgvSettings.DragOver += (s, e) => {
                    e.Effect = DragDropEffects.Move;
                    Point p = dgvSettings.PointToClient(new Point(e.X, e.Y));
                    var hit = dgvSettings.HitTest(p.X, p.Y);
                    int newToIdx = hit.RowIndex;
                    if (newToIdx < 0) newToIdx = dgvSettings.Rows.Count - 1;

                    if (dragToIdx != newToIdx) {
                        dragToIdx = newToIdx;
                        dgvSettings.Invalidate(); 
                    }
                };

                dgvSettings.DragDrop += (s, e) => {
                    if (e.Data.GetDataPresent(typeof(DataGridViewRow))) {
                        Point p = dgvSettings.PointToClient(new Point(e.X, e.Y));
                        var hit = dgvSettings.HitTest(p.X, p.Y);
                        int targetIdx = hit.RowIndex;
                        if (targetIdx < 0) targetIdx = dgvSettings.Rows.Count - 1;

                        if (dragFromIdx >= 0 && dragFromIdx != targetIdx) {
                            dgvSettings.EndEdit();
                            DataGridViewRow rowToMove = dgvSettings.Rows[dragFromIdx];
                            dgvSettings.Rows.RemoveAt(dragFromIdx);
                            dgvSettings.Rows.Insert(targetIdx, rowToMove);
                            dgvSettings.ClearSelection();
                            dgvSettings.Rows[targetIdx].Selected = true;
                        }
                    }
                    dragToIdx = -1;
                    dgvSettings.Invalidate();
                };

                dgvSettings.Paint += (s, e) => {
                    if (dragToIdx >= 0 && dragToIdx < dgvSettings.Rows.Count) {
                        Rectangle r = dgvSettings.GetRowDisplayRectangle(dragToIdx, false);
                        using (Pen pen = new Pen(Color.Red, 3)) {
                            e.Graphics.DrawLine(pen, r.Left, r.Top, r.Right, r.Top);
                        }
                    }
                };

                dgvSettings.CellContentClick += (s, e) => {
                    if (e.ColumnIndex == 0 && e.RowIndex >= 0) {
                        dgvSettings.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    }
                };

                Panel pnlRight = new Panel { Dock = DockStyle.Right, Width = 110, Padding = new Padding(10, 0, 10, 0) };
                Button btnUp = new Button { Text = "↑ 上移", Dock = DockStyle.Top, Height = 50, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand };
                Button btnDown = new Button { Text = "↓ 下移", Dock = DockStyle.Top, Height = 50, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, Margin = new Padding(0, 10, 0, 0) };
                
                btnUp.Click += (s, e) => {
                    if (dgvSettings.SelectedRows.Count > 0) {
                        int idx = dgvSettings.SelectedRows[0].Index;
                        if (idx > 0) {
                            dgvSettings.EndEdit();
                            var row = dgvSettings.Rows[idx];
                            dgvSettings.Rows.RemoveAt(idx);
                            dgvSettings.Rows.Insert(idx - 1, row);
                            dgvSettings.ClearSelection();
                            dgvSettings.Rows[idx - 1].Selected = true;
                        }
                    }
                };

                btnDown.Click += (s, e) => {
                    if (dgvSettings.SelectedRows.Count > 0) {
                        int idx = dgvSettings.SelectedRows[0].Index;
                        if (idx < dgvSettings.Rows.Count - 1) {
                            dgvSettings.EndEdit();
                            var row = dgvSettings.Rows[idx];
                            dgvSettings.Rows.RemoveAt(idx);
                            dgvSettings.Rows.Insert(idx + 1, row);
                            dgvSettings.ClearSelection();
                            dgvSettings.Rows[idx + 1].Selected = true;
                        }
                    }
                };

                pnlRight.Controls.Add(btnDown);
                pnlRight.Controls.Add(new Panel { Height = 10, Dock = DockStyle.Top }); 
                pnlRight.Controls.Add(btnUp);

                Panel pnlCenter = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
                pnlCenter.Controls.Add(dgvSettings);
                pnlCenter.Controls.Add(pnlRight);

                f.Controls.Add(pnlCenter);
                
                // 將 lblTop 推到最上方，這可以確保 pnlCenter 在其下
                lblTop.SendToBack();

                Button btnSaveLocal = new Button { Text = "💾 儲存並套用設定", Dock = DockStyle.Bottom, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSaveLocal.FlatAppearance.BorderSize = 0;

                btnSaveLocal.Click += delegate(object s, EventArgs ev) { 
                    dgvSettings.EndEdit();
                    DataManager.ClearGridConfig("AuditDash", gridId, "Visibility");
                    DataManager.ClearGridConfig("AuditDash", gridId, "Order");

                    List<string> orderedCols = new List<string>();

                    for (int i = 0; i < dgvSettings.Rows.Count; i++) {
                        string colName = dgvSettings.Rows[i].Cells["ColName"].Value.ToString(); 
                        bool isChecked = Convert.ToBoolean(dgvSettings.Rows[i].Cells["Visible"].Value);
                        
                        orderedCols.Add(colName);
                        
                        if (targetDgv.Columns.Contains(colName)) {
                            targetDgv.Columns[colName].Visible = isChecked;
                            targetDgv.Columns[colName].DisplayIndex = i; 
                        }
                        
                        DataManager.SaveGridConfig("AuditDash", gridId, "Visibility", colName, isChecked ? "1" : "0");
                    } 

                    DataManager.SaveGridConfig("AuditDash", gridId, "Order", "All", string.Join(",", orderedCols));

                    f.DialogResult = DialogResult.OK; 
                };
                
                f.Controls.Add(btnSaveLocal); 
                f.ShowDialog();
            }
        }

        // ==========================================
        // 匯出 Excel 系統
        // ==========================================
        private void ExportToExcel()
        {
            if ((_dgvResultA.DataSource == null || _dgvResultA.Rows.Count == 0) &&
                (_dgvResultB.DataSource == null || _dgvResultB.Rows.Count == 0))
            {
                MessageBox.Show("目前沒有任何查詢結果可供匯出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "稽核資料查詢結果_" + DateTime.Now.ToString("yyyyMMdd_HHmm") })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        using (ExcelPackage p = new ExcelPackage())
                        {
                            if (_dgvResultA.DataSource != null && _dgvResultA.Rows.Count > 0)
                            {
                                var wsA = p.Workbook.Worksheets.Add("巡檢缺失");
                                DataTable dtA = GetVisibleData(_dgvResultA);
                                wsA.Cells["A1"].LoadFromDataTable(dtA, true);
                                FormatExcelSheet(wsA, dtA, _dgvResultA);
                            }

                            if (_dgvResultB.DataSource != null && _dgvResultB.Rows.Count > 0)
                            {
                                var wsB = p.Workbook.Worksheets.Add("追蹤改善");
                                DataTable dtB = GetVisibleData(_dgvResultB);
                                wsB.Cells["A1"].LoadFromDataTable(dtB, true);
                                FormatExcelSheet(wsB, dtB, _dgvResultB);
                            }

                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        
                        MessageBox.Show("Excel 報表匯出成功！已套用完美排版。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void FormatExcelSheet(ExcelWorksheet ws, DataTable dt, DataGridView sourceDgv)
        {
            using (var range = ws.Cells[1, 1, 1, dt.Columns.Count]) {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }

            var dataRange = ws.Cells[2, 1, dt.Rows.Count + 1, dt.Columns.Count];
            dataRange.Style.WrapText = true;
            dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            
            for (int c = 1; c <= dt.Columns.Count; c++) {
                string colName = dt.Columns[c - 1].ColumnName;
                
                // 參考 UI 的欄寬
                if (sourceDgv.Columns.Contains(colName)) {
                    float uiWidth = sourceDgv.Columns[colName].Width;
                    float excelWidth = (uiWidth / 7f); 
                    ws.Column(c).Width = excelWidth < 10 ? 10 : excelWidth;
                }
            }

            Font excelFont = new Font("Microsoft JhengHei UI", 11F);
            
            using (Bitmap dummyBmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(dummyBmp))
            {
                for (int r = 0; r < dt.Rows.Count; r++) 
                {
                    float maxRowHeightPts = 25f; 

                    for (int c = 0; c < dt.Columns.Count; c++) 
                    {
                        string text = dt.Rows[r][c]?.ToString() ?? "";
                        if (string.IsNullOrEmpty(text)) continue;

                        double colWidthChars = ws.Column(c + 1).Width;
                        int pixelWidth = (int)(colWidthChars * 7.0); 

                        SizeF sz = g.MeasureString(text, excelFont, pixelWidth);
                        float ptHeight = sz.Height * 0.75f;

                        if (ptHeight > maxRowHeightPts) {
                            maxRowHeightPts = ptHeight;
                        }
                    }

                    ws.Row(r + 2).Height = maxRowHeightPts > 25f ? maxRowHeightPts + 5f : 25f;
                    ws.Row(r + 2).CustomHeight = true;
                }
                
                ws.Row(1).Height = 25f;
                ws.Row(1).CustomHeight = true;
            }
        }

        private DataTable GetVisibleData(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            
            var visCols = dgv.Columns.Cast<DataGridViewColumn>()
                                     .Where(c => c.Visible)
                                     .OrderBy(c => c.DisplayIndex)
                                     .ToList();
            
            foreach (var col in visCols) {
                dt.Columns.Add(col.HeaderText.Replace("\n", ""));
            }

            foreach (DataGridViewRow row in dgv.Rows) {
                if (row.IsNewRow) continue;
                DataRow dRow = dt.NewRow();
                for (int i = 0; i < visCols.Count; i++) {
                    var cellVal = row.Cells[visCols[i].Index].Value;
                    dRow[i] = cellVal != null ? cellVal.ToString() : "";
                }
                dt.Rows.Add(dRow);
            }
            return dt;
        }
    }
}
