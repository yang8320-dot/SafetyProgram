/// FILE: Safety_System/Dashboard/App_AuditDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

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

        // UI 控制項 - 結果區 A (日期區間)
        private Label _lblResultATitle;
        private DataGridView _dgvResultA;
        private Button _btnSettingsA;

        // UI 控制項 - 結果區 B (單號追蹤)
        private Label _lblResultBTitle;
        private DataGridView _dgvResultB;
        private Button _btnSettingsB;

        // 資料庫快取
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 下拉選單對映模型
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
            Label lblTitle = new Label { Text = "🛡️ 稽核資料查詢", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
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

            // 多行文字框與查詢按鈕
            FlowLayoutPanel flpText = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 5, 0, 5) };
            _txtFormNumbers = new TextBox { Multiline = true, Width = 350, Height = 100, Font = new Font("Consolas", 12F), ScrollBars = ScrollBars.Vertical, Margin = new Padding(5, 0, 20, 0) };
            
            _btnSearch = new Button { Text = "🚀 執行雙向查詢", Size = new Size(200, 60), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 20, 0, 0) };
            _btnSearch.FlatAppearance.BorderSize = 0;
            _btnSearch.Click += async (s, e) => await ExecuteSearchAsync();

            flpText.Controls.AddRange(new Control[] {
                new Label { Text = "查詢巡檢表單單號:\n(每行代表一段查詢文字)", AutoSize = true, Margin = new Padding(30, 5, 0, 0), Font = new Font("Microsoft JhengHei UI", 11F), ForeColor = Color.DimGray },
                _txtFormNumbers,
                _btnSearch
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
            
            Panel pnlHeaderA = new Panel { Dock = DockStyle.Top, Height = 40 };
            _lblResultATitle = new Label { Text = "查詢區間：尚未查詢", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Left, AutoSize = true };
            _btnSettingsA = new Button { Text = "⚙️ 顯示設定", Size = new Size(120, 32), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Right, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSettingsA.FlatAppearance.BorderSize = 0;
            _btnSettingsA.Click += (s, e) => OpenGridSettings(_dgvResultA, "GridA");

            pnlHeaderA.Controls.Add(_lblResultATitle);
            pnlHeaderA.Controls.Add(_btnSettingsA);

            _dgvResultA = CreateStandardGrid();
            
            boxResultA.Controls.Add(_dgvResultA);
            boxResultA.Controls.Add(pnlHeaderA);
            masterLayout.Controls.Add(boxResultA, 0, 2);

            // ==========================================
            // 第四個框：取得查詢資料 (B. 單號追蹤)
            // ==========================================
            GroupBox boxResultB = new GroupBox { Text = "📑 查詢資料 (依單號追蹤)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0, 0, 0, 20) };
            
            Panel pnlHeaderB = new Panel { Dock = DockStyle.Top, Height = 40 };
            _lblResultBTitle = new Label { Text = "巡檢異常改善單 - 追蹤", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.Chocolate, Dock = DockStyle.Left, AutoSize = true };
            _btnSettingsB = new Button { Text = "⚙️ 顯示設定", Size = new Size(120, 32), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Dock = DockStyle.Right, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSettingsB.FlatAppearance.BorderSize = 0;
            _btnSettingsB.Click += (s, e) => OpenGridSettings(_dgvResultB, "GridB");

            pnlHeaderB.Controls.Add(_lblResultBTitle);
            pnlHeaderB.Controls.Add(_btnSettingsB);

            _dgvResultB = CreateStandardGrid();
            
            boxResultB.Controls.Add(_dgvResultB);
            boxResultB.Controls.Add(pnlHeaderB);
            masterLayout.Controls.Add(boxResultB, 0, 3);

            mainScrollPanel.Controls.Add(masterLayout);

            InitDropdownLogic();

            return mainScrollPanel;
        }

        private DataGridView CreateStandardGrid()
        {
            DataGridView dgv = new DataGridView { 
                Dock = DockStyle.Top, 
                Height = 350,
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                AllowUserToDeleteRows = false, 
                ReadOnly = true, 
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                Font = new Font("Microsoft JhengHei UI", 11F), 
                BorderStyle = BorderStyle.FixedSingle, 
                Margin = new Padding(0, 10, 0, 0)
            };
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 40;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            // 啟用雙重緩衝避免閃爍
            typeof(DataGridView).InvokeMember("DoubleBuffered", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, 
                null, dgv, new object[] { true });

            return dgv;
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
                
                // 智慧預選：如果欄位中有「單號」相關字眼，自動選取
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

            // 智慧預設選中工安庫的巡檢紀錄
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

        private async System.Threading.Tasks.Task ExecuteSearchAsync()
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

            await System.Threading.Tasks.Task.Run(() => {
                // 取得目標表所有資料
                rawData = DataManager.GetTableData(dbName, tbName, "", "", "");
                if (rawData == null || rawData.Rows.Count == 0) return;

                // 處理 A 區塊 (日期篩選)
                dvA = new DataView(rawData);
                string dateCol = GetDateColumn(rawData);
                if (!string.IsNullOrEmpty(dateCol)) {
                    dvA.RowFilter = $"[{dateCol}] >= '{sDate:yyyy-MM-dd}' AND [{dateCol}] <= '{eDate:yyyy-MM-dd}'";
                }

                // 處理 B 區塊 (單號篩選)
                dvB = new DataView(rawData);
                string[] forms = _txtFormNumbers.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (forms.Length > 0 && !string.IsNullOrEmpty(colName) && rawData.Columns.Contains(colName)) {
                    string inClause = string.Join(",", forms.Select(f => $"'{f.Trim().Replace("'", "''")}'"));
                    dvB.RowFilter = $"[{colName}] IN ({inClause})";
                } else {
                    dvB.RowFilter = "1=0"; // 沒輸入單號則顯示空表
                }
            });

            if (rawData != null) {
                _dgvResultA.DataSource = dvA.ToTable();
                _dgvResultB.DataSource = dvB.ToTable();

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
        // 欄位顯示設定系統 (利用既有的 GridConfigs)
        // ==========================================
        private void ApplyGridSettings(DataGridView dgv, string gridId)
        {
            if (dgv.Columns.Count == 0) return;

            var visibilityDict = DataManager.LoadGridConfig("AuditDash", gridId, "Visibility");

            foreach (DataGridViewColumn col in dgv.Columns) {
                if (col.Name == "Id") {
                    col.Visible = false; continue;
                }

                if (visibilityDict.ContainsKey(col.Name)) {
                    col.Visible = (visibilityDict[col.Name] == "1");
                } else {
                    // 預設全顯示，除了系統欄位
                    if (col.Name == "最後修改人" || col.Name == "修改時間") col.Visible = false;
                    else col.Visible = true;
                }
            }
        }

        private void OpenGridSettings(DataGridView dgv, string gridId)
        {
            if (dgv.Columns.Count == 0) {
                MessageBox.Show("請先執行查詢產生資料後，再進行欄位設定。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (Form f = new Form { Text = "👁️ 顯示欄位設定", Size = new Size(350, 500), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White }) 
            {
                Label lblTop = new Label { Text = "請勾選欲顯示在表格中的欄位：", Dock = DockStyle.Top, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.SteelBlue }; 
                f.Controls.Add(lblTop);
                
                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.None, Padding = new Padding(10), BackColor = Color.WhiteSmoke };
                
                foreach (DataGridViewColumn col in dgv.Columns) { 
                    if (col.Name == "Id") continue; 
                    clbCols.Items.Add(col.Name, col.Visible); 
                }
                f.Controls.Add(clbCols);
                
                Button btnSaveLocal = new Button { Text = "💾 儲存並套用設定", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnSaveLocal.FlatAppearance.BorderSize = 0;

                btnSaveLocal.Click += delegate(object s, EventArgs ev) { 
                    DataManager.ClearGridConfig("AuditDash", gridId, "Visibility");

                    for (int i = 0; i < clbCols.Items.Count; i++) { 
                        string colName = clbCols.Items[i].ToString(); 
                        bool isChecked = clbCols.GetItemChecked(i); 
                        
                        if (dgv.Columns.Contains(colName)) dgv.Columns[colName].Visible = isChecked; 
                        
                        DataManager.SaveGridConfig("AuditDash", gridId, "Visibility", colName, isChecked ? "1" : "0");
                    } 
                    f.DialogResult = DialogResult.OK; 
                };
                
                f.Controls.Add(btnSaveLocal); 
                f.ShowDialog();
            }
        }
    }
}
