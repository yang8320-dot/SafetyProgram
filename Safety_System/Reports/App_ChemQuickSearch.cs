/// FILE: Safety_System/Reports/App_ChemQuickSearch.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemQuickSearch
    {
        private TextBox _txtName;
        private TextBox _txtCAS;
        private Button _btnSearch;
        private Button _btnSettings;
        private Label _lblStatus;
        private FlowLayoutPanel _flpResultsContainer; 
        
        private const string DbName = "Chemical";
        
        // 🟢 替換：移除 TXT 路徑，改用 SQLite
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        // 確保系統未初始化時能預先建表，避免設定欄位時找不到表
        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            { "EnvTesting", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [分類] TEXT, [中文名稱] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ExposureLimits", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [分類] TEXT, [種類] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [容許濃度ppm] TEXT, [容許濃度mgm3] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ToxicSubstances", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [分類] TEXT, [種類] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [管制濃度百分比] TEXT, [分級運作量kg] TEXT, [毒性分類] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ConcernedChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [管制濃度百分比] TEXT, [管制行為] TEXT, [具有危害性之關注化學物質註記] TEXT, [分級運作量kg] TEXT, [定期申報頻率] TEXT, [毒性分類] TEXT, [包裝容器規定] TEXT, [記錄] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PriorityMgmtChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [CASNO] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ControlledChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SpecificChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [類別] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "OrganicSolvents", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [類別] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkerHealthProtect", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PublicHazardous", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [分類] TEXT, [名稱] TEXT, [種類] TEXT, [分級] TEXT, [管制量] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "AirPollutionEmerg", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [CASNO] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FactoryHazardous", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [分類] TEXT, [名稱] TEXT, [種類] TEXT, [管制量] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" }
        };

        private class ChemTableInfo {
            public string TableName;
            public string Title;
            public string ExtraNotice;
            public GroupBox GBox;
            public DataGridView Dgv;
            public DataTable ResultData;
            public List<string> VisibleColumns; 
        }

        private List<ChemTableInfo> _tableInfos = new List<ChemTableInfo> {
            new ChemTableInfo { TableName="EnvTesting", Title="1. 環測項目", ExtraNotice="" },
            new ChemTableInfo { TableName="ExposureLimits", Title="2. 勞工暴露容許濃度", ExtraNotice="" },
            new ChemTableInfo { TableName="ToxicSubstances", Title="3. 毒性物質", ExtraNotice="" },
            new ChemTableInfo { TableName="ConcernedChem", Title="4. 關注性化學物質", ExtraNotice="" },
            new ChemTableInfo { TableName="PriorityMgmtChem", Title="5. 優先管理化學品", ExtraNotice="" },
            new ChemTableInfo { TableName="ControlledChem", Title="6. 管制化學品", ExtraNotice="" },
            new ChemTableInfo { TableName="SpecificChem", Title="7. 特定化學物質", ExtraNotice="需設置【特化主管】" },
            new ChemTableInfo { TableName="OrganicSolvents", Title="8. 有機溶劑", ExtraNotice="需設置【有機溶劑作業主管】" },
            new ChemTableInfo { TableName="WorkerHealthProtect", Title="9. 勞工健康保護", ExtraNotice="需【特殊體檢】" },
            new ChemTableInfo { TableName="PublicHazardous", Title="10. 公共危險物品", ExtraNotice="" },
            new ChemTableInfo { TableName="AirPollutionEmerg", Title="11. 空污緊急應變", ExtraNotice="" },
            new ChemTableInfo { TableName="FactoryHazardous", Title="12. 工廠危險物品申報", ExtraNotice="" }
        };

        // 🟢 新增：資料庫初始化邏輯
        private void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS [ChemQuickSearchVisibility] (
                        ItemKey TEXT PRIMARY KEY, IsVisible INTEGER);";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        public Control GetView()
        {
            InitDatabase(); // 🟢 初始化 SQLite 表格
            
            foreach (var info in _tableInfos) {
                if (_schemaMap.ContainsKey(info.TableName)) {
                    DataManager.InitTable(DbName, info.TableName, $"CREATE TABLE IF NOT EXISTS [{info.TableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {_schemaMap[info.TableName]});");
                }
            }

            LoadVisibilitySettings();

            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                Padding = new Padding(20), 
                RowCount = 2,
                BackColor = Color.WhiteSmoke
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            FlowLayoutPanel pnlAction = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 10),
                WrapContents = false 
            };
            
            Button btnPdf = new Button { 
                Text = "📄 導出PDF", 
                Size = new Size(130, 40), 
                BackColor = Color.IndianRed, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 5, 10, 5)
            };
            btnPdf.Click += (s, e) => ExportToPdf();

            _btnSettings = new Button {
                Text = "⚙️ 顯示欄位設定",
                Size = new Size(180, 40),
                BackColor = Color.LightSlateGray,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 5, 10, 5)
            };
            _btnSettings.Click += (s, e) => OpenSettingsDialog();

            _btnSearch = new Button {
                Text = "🚀 開始執行交叉檢索",
                Size = new Size(220, 40),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5, 5, 15, 5) 
            };
            _btnSearch.Click += async (s, e) => await ExecuteSearchAsync();

            _lblStatus = new Label {
                Text = "準備就緒。請輸入條件後點擊查詢或按下 Enter 鍵。",
                ForeColor = Color.DimGray,
                Font = new Font("Microsoft JhengHei UI", 11F),
                AutoSize = true,
                Margin = new Padding(0, 15, 0, 0) 
            };

            pnlAction.Controls.AddRange(new Control[] { btnPdf, _btnSettings, _btnSearch, _lblStatus });
            mainLayout.Controls.Add(pnlAction, 0, 0);

            GroupBox boxMain = new GroupBox { 
                Text = "🔍 化學品法規符核度查詢", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                ForeColor = Color.DarkCyan, 
                Padding = new Padding(15) 
            };
            
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));  
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  

            Panel pnlSearch = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            pnlSearch.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, pnlSearch.ClientRectangle, Color.FromArgb(200, 200, 200), ButtonBorderStyle.Solid);
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(45, 62, 80))) {
                    e.Graphics.FillRectangle(brush, 0, 0, 6, pnlSearch.Height); 
                }
            };

            FlowLayoutPanel flpSearch = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(15, 18, 15, 15) };
            
            Label lbl1 = new Label { Text = "化學品名稱關鍵字：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 5, 5, 0) };
            _txtName = new TextBox { Width = 250, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtName.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) { e.Handled = true; e.SuppressKeyPress = true; await ExecuteSearchAsync(); } };
            
            Label lbl2 = new Label { Text = "CAS No. 編號：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(40, 5, 5, 0) };
            _txtCAS = new TextBox { Width = 200, Font = new Font("Microsoft JhengHei UI", 13F) };
            _txtCAS.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) { e.Handled = true; e.SuppressKeyPress = true; await ExecuteSearchAsync(); } };

            flpSearch.Controls.AddRange(new Control[] { lbl1, _txtName, lbl2, _txtCAS });
            pnlSearch.Controls.Add(flpSearch);
            innerTable.Controls.Add(pnlSearch, 0, 0);

            GroupBox sub4 = new GroupBox { 
                Text = "📊 檢索結果明細 (查無資料之分類將自動隱藏)", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), 
                Margin = new Padding(0, 10, 0, 0) 
            };

            _flpResultsContainer = new FlowLayoutPanel {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(5, 5, 30, 5), 
                BackColor = Color.White
            };
            
            _flpResultsContainer.Resize += (s, e) => {
                foreach (Control c in _flpResultsContainer.Controls) {
                    if (c is GroupBox gb) gb.Width = _flpResultsContainer.ClientSize.Width - 30;
                }
            };
            
            foreach (var info in _tableInfos) {
                info.GBox = new GroupBox {
                    Text = info.Title + (string.IsNullOrEmpty(info.ExtraNotice) ? "" : " - " + info.ExtraNotice),
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                    ForeColor = string.IsNullOrEmpty(info.ExtraNotice) ? Color.DarkSlateBlue : Color.Crimson,
                    AutoSize = false, 
                    Margin = new Padding(0, 0, 0, 10), 
                    Padding = new Padding(5, 20, 5, 10), 
                    Visible = false 
                };

                info.Dgv = new DataGridView { 
                    Dock = DockStyle.Fill, 
                    BackgroundColor = Color.White, 
                    AllowUserToAddRows = false, 
                    ReadOnly = true, 
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect, 
                    RowHeadersVisible = false,
                    BorderStyle = BorderStyle.None,
                    ScrollBars = ScrollBars.None, 
                    Font = new Font("Microsoft JhengHei UI", 11F),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                };
                info.Dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True; 
                
                info.Dgv.EnableHeadersVisualStyles = false;
                info.Dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
                info.Dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                info.Dgv.ColumnHeadersHeight = 40; 
                
                info.Dgv.DefaultCellStyle.BackColor = Color.White;
                info.Dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
                info.Dgv.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue; 
                info.Dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

                info.GBox.Controls.Add(info.Dgv);
                _flpResultsContainer.Controls.Add(info.GBox);
            }

            sub4.Controls.Add(_flpResultsContainer);
            innerTable.Controls.Add(sub4, 0, 1);
            
            boxMain.Controls.Add(innerTable);
            mainLayout.Controls.Add(boxMain, 0, 1);

            return mainLayout;
        }

        private string EscapeLikeValue(string valueWithoutWildcards)
        {
            if (string.IsNullOrEmpty(valueWithoutWildcards)) return "";
            
            return valueWithoutWildcards
                .Replace("[", "[[]")
                .Replace("]", "[]]")
                .Replace("%", "[%]")
                .Replace("*", "[*]")
                .Replace("'", "''");
        }

        private async Task ExecuteSearchAsync()
        {
            string nameKey = _txtName.Text.Trim();
            string casKey = _txtCAS.Text.Trim();

            if (string.IsNullOrEmpty(nameKey) && string.IsNullOrEmpty(casKey)) {
                MessageBox.Show("請至少輸入一個查詢關鍵字或 CAS No。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _btnSearch.Enabled = false;
            _btnSearch.Text = "⏳ 檢索與排版中...";
            _lblStatus.Text = "正在背景非同步檢索資料庫，請稍候...";
            _lblStatus.ForeColor = Color.OrangeRed;

            string safeNameKey = EscapeLikeValue(nameKey);
            string safeCasKey = EscapeLikeValue(casKey);

            await Task.Run(() => {
                foreach (var info in _tableInfos) {
                    try {
                        DataTable dt = DataManager.GetTableData(DbName, info.TableName, "", "", "");
                        if (dt != null && dt.Rows.Count > 0) {
                            DataView dv = dt.DefaultView;
                            List<string> matchConditions = new List<string>();

                            // 1. 名稱模糊檢索
                            if (!string.IsNullOrEmpty(safeNameKey)) {
                                List<string> nameFilters = new List<string>();
                                string[] possibleNameCols = { "中文名稱", "英文名稱", "名稱", "種類", "化學品名稱", "內容", "危害成份" };
                                foreach(var c in possibleNameCols) {
                                    if (dt.Columns.Contains(c)) {
                                        nameFilters.Add($"[{c}] LIKE '%{safeNameKey}%'");
                                    }
                                }
                                if (nameFilters.Count > 0) {
                                    matchConditions.Add("(" + string.Join(" OR ", nameFilters) + ")");
                                } else {
                                    matchConditions.Add("1=0"); 
                                }
                            }

                            // 2. CAS Number 檢索 
                            if (!string.IsNullOrEmpty(safeCasKey)) {
                                List<string> casFilters = new List<string>();
                                string[] possibleCasCols = { "CASNO", "CAS_No", "化學式" };
                                foreach(var c in possibleCasCols) {
                                    if (dt.Columns.Contains(c)) {
                                        casFilters.Add($"[{c}] LIKE '%{safeCasKey}%'");
                                    }
                                }
                                
                                if (casFilters.Count > 0) {
                                    matchConditions.Add("(" + string.Join(" OR ", casFilters) + ")");
                                } else if (string.IsNullOrEmpty(safeNameKey)) {
                                    matchConditions.Add("1=0");
                                }
                            }

                            dv.RowFilter = matchConditions.Count > 0 ? string.Join(" AND ", matchConditions) : "1=0";
                            info.ResultData = dv.ToTable();

                            info.VisibleColumns = new List<string>();
                            if (info.ResultData.Rows.Count > 0) {
                                foreach (DataColumn col in info.ResultData.Columns) {
                                    if (col.ColumnName == "Id") continue;
                                    
                                    string dictKey = $"{info.TableName}_{col.ColumnName}";

                                    bool isAllowedToShow = true;
                                    if (_columnVisibility.ContainsKey(dictKey)) {
                                        isAllowedToShow = _columnVisibility[dictKey];
                                    } else {
                                        if (col.ColumnName == "最後修改人" || col.ColumnName == "修改時間" || col.ColumnName.Contains("時間") || col.ColumnName.Contains("修改人")) {
                                            isAllowedToShow = false;
                                        }
                                    }

                                    if (!isAllowedToShow) continue;

                                    bool hasValue = false;
                                    foreach (DataRow row in info.ResultData.Rows) {
                                        if (row[col] != DBNull.Value && !string.IsNullOrWhiteSpace(row[col].ToString())) {
                                            hasValue = true; break;
                                        }
                                    }
                                    if (hasValue) info.VisibleColumns.Add(col.ColumnName);
                                }
                            }
                        } else {
                            info.ResultData = null;
                        }
                    } catch {
                        info.ResultData = null;
                    }
                }
            });

            _flpResultsContainer.SuspendLayout();
            int totalFound = 0;

            foreach (var info in _tableInfos) {
                if (info.ResultData != null && info.ResultData.Rows.Count > 0) {
                    
                    info.Dgv.DataSource = info.ResultData;

                    foreach (DataGridViewColumn col in info.Dgv.Columns) {
                        col.Visible = info.VisibleColumns.Contains(col.Name);
                    }

                    info.GBox.Width = _flpResultsContainer.ClientSize.Width - 30;
                    info.Dgv.AutoResizeRows(); 

                    int exactGridHeight = info.Dgv.ColumnHeadersHeight;
                    foreach(DataGridViewRow r in info.Dgv.Rows) {
                        exactGridHeight += r.Height;
                    }
                    info.GBox.Height = info.GBox.Padding.Top + exactGridHeight + 50; 
                    
                    info.Dgv.ClearSelection();
                    info.GBox.Visible = true;
                    totalFound += info.ResultData.Rows.Count;

                    await Task.Delay(5); 
                } else {
                    info.GBox.Visible = false;
                    info.Dgv.DataSource = null; 
                }
            }

            _flpResultsContainer.ResumeLayout(true);
            _btnSearch.Enabled = true;
            _btnSearch.Text = "🚀 開始執行交叉檢索";
            _lblStatus.Text = $"檢索完成！共在各分類中找到 {totalFound} 筆資料。";
            _lblStatus.ForeColor = Color.ForestGreen;

            if (totalFound == 0) {
                MessageBox.Show("於所有法規資料庫中皆查無符合條件之項目。", "檢索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 🟢 替換為 SQLite
        private void LoadVisibilitySettings()
        {
            _columnVisibility.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM ChemQuickSearchVisibility", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            _columnVisibility[reader["ItemKey"].ToString()] = Convert.ToInt32(reader["IsVisible"]) == 1;
                        }
                    }
                }
            } catch { }
        }

        // 🟢 替換為 SQLite
        private void SaveVisibilitySettings()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        new SQLiteCommand("DELETE FROM ChemQuickSearchVisibility", conn, trans).ExecuteNonQuery();
                        
                        string sql = "INSERT INTO ChemQuickSearchVisibility (ItemKey, IsVisible) VALUES (@K, @V)";
                        foreach (var kvp in _columnVisibility) {
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@K", kvp.Key);
                                cmd.Parameters.AddWithValue("@V", kvp.Value ? 1 : 0);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
            } catch { }
        }

        private void OpenSettingsDialog()
        {
            using (Form f = new Form { Text = "⚙️ 顯示欄位設定", Size = new Size(650, 550), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, BackColor = Color.WhiteSmoke }) 
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));

                Label lblTop = new Label { 
                    Text = "請選擇分類並勾選查詢時【允許顯示】的欄位：", 
                    Dock = DockStyle.Fill, 
                    Padding = new Padding(10), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = Color.SteelBlue, 
                    AutoSize = true 
                };
                tlp.Controls.Add(lblTop, 0, 0);

                SplitContainer split = new SplitContainer { 
                    Dock = DockStyle.Fill, 
                    SplitterDistance = 250, 
                    FixedPanel = FixedPanel.Panel1, 
                    Padding = new Padding(10) 
                };

                ListBox lbTables = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
                foreach (var info in _tableInfos) lbTables.Items.Add(info.Title);
                split.Panel1.Controls.Add(lbTables);

                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle };
                split.Panel2.Controls.Add(clbCols);
                tlp.Controls.Add(split, 0, 1);

                Button btnSave = new Button { 
                    Text = "💾 儲存並關閉", 
                    Dock = DockStyle.Fill, 
                    BackColor = Color.ForestGreen, 
                    ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    Cursor = Cursors.Hand, 
                    FlatStyle = FlatStyle.Flat 
                };
                tlp.Controls.Add(btnSave, 0, 2);

                f.Controls.Add(tlp);

                lbTables.SelectedIndexChanged += (s, e) => {
                    if (lbTables.SelectedIndex < 0) return;
                    clbCols.Items.Clear();
                    string tblName = _tableInfos[lbTables.SelectedIndex].TableName;
                    var cols = DataManager.GetColumnNames(DbName, tblName);
                    
                    if (!cols.Contains("最後修改人")) cols.Add("最後修改人");
                    if (!cols.Contains("修改時間")) cols.Add("修改時間");

                    foreach (var c in cols) {
                        if (c == "Id") continue;
                        string key = $"{tblName}_{c}";
                        
                        bool isChecked = true;
                        if (_columnVisibility.ContainsKey(key)) {
                            isChecked = _columnVisibility[key];
                        } else {
                            if (c == "最後修改人" || c == "修改時間" || c.Contains("時間") || c.Contains("修改人")) {
                                isChecked = false;
                            }
                        }
                        
                        clbCols.Items.Add(c, isChecked);
                    }
                };

                clbCols.ItemCheck += (s, e) => {
                    string colName = clbCols.Items[e.Index].ToString();

                    if ((colName == "最後修改人" || colName == "修改時間" || colName.Contains("時間") || colName.Contains("修改人")) && e.NewValue == CheckState.Checked) {
                        if (!AuthManager.VerifyLv3Only("顯示系統修改紀錄需要最高權限\n請輸入【Lv3系統管理者】密碼：")) {
                            e.NewValue = CheckState.Unchecked;
                            return; 
                        }
                    }

                    if (lbTables.SelectedIndex < 0) return;
                    string tblName = _tableInfos[lbTables.SelectedIndex].TableName;
                    _columnVisibility[$"{tblName}_{colName}"] = e.NewValue == CheckState.Checked;
                };

                btnSave.Click += (s, e) => {
                    SaveVisibilitySettings();
                    f.DialogResult = DialogResult.OK;
                };

                if (lbTables.Items.Count > 0) lbTables.SelectedIndex = 0;

                if (f.ShowDialog() == DialogResult.OK && (!string.IsNullOrEmpty(_txtName.Text) || !string.IsNullOrEmpty(_txtCAS.Text))) {
                    _btnSearch.PerformClick();
                }
            }
        }

        private void ExportToPdf()
        {
            var visibleTables = _tableInfos.Where(t => t.GBox.Visible && t.ResultData != null && t.ResultData.Rows.Count > 0).ToList();
            if (visibleTables.Count == 0) {
                MessageBox.Show("目前暫無搜尋數據可供導出。"); return;
            }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = false; 
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 40, 40);
            
            int currentTableIndex = 0;
            int currentRowIndex = 0;

            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                Font fTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                Font fSubTitle = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                g.DrawString("化學品法規符核度查詢報表", fTitle, Brushes.Black, x, y);
                y += 35;
                g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   台灣玻璃彰濱廠", fBody, Brushes.Gray, x, y);
                y += 35;

                while (currentTableIndex < visibleTables.Count) {
                    var info = visibleTables[currentTableIndex];
                    var visCols = info.Dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                    
                    if (visCols.Count == 0) { currentTableIndex++; continue; }

                    float totalW = visCols.Sum(c => c.Width);
                    float[] colWidths = new float[visCols.Count];
                    for (int i = 0; i < visCols.Count; i++) {
                        colWidths[i] = (visCols[i].Width / totalW) * e.MarginBounds.Width;
                    }

                    if (currentRowIndex == 0) {
                        if (y + 80 > e.MarginBounds.Bottom) { e.HasMorePages = true; return; }
                        
                        g.DrawString(info.Title + " " + info.ExtraNotice, fSubTitle, Brushes.DarkSlateBlue, x, y);
                        y += 25;

                        float currX = x;
                        for (int i = 0; i < visCols.Count; i++) {
                            RectangleF rect = new RectangleF(currX, y, colWidths[i], 35);
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            g.DrawString(visCols[i].HeaderText, fHead, Brushes.Black, rect, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                            currX += colWidths[i];
                        }
                        y += 35;
                    }

                    StringFormat fmtWrap = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
                    
                    while (currentRowIndex < info.ResultData.Rows.Count) {
                        float maxRowHeight = 30; 
                        for (int i = 0; i < visCols.Count; i++) {
                            string val = info.Dgv[visCols[i].Index, currentRowIndex].Value?.ToString() ?? "";
                            SizeF sSize = g.MeasureString(val, fBody, (int)colWidths[i], fmtWrap);
                            if (sSize.Height + 10 > maxRowHeight) maxRowHeight = sSize.Height + 10;
                        }

                        if (y + maxRowHeight > e.MarginBounds.Bottom) { 
                            e.HasMorePages = true; 
                            return; 
                        }

                        float currX = x;
                        for (int i = 0; i < visCols.Count; i++) {
                            RectangleF rect = new RectangleF(currX, y, colWidths[i], maxRowHeight);
                            g.DrawRectangle(Pens.Black, rect.X, rect.Y, rect.Width, rect.Height);
                            string val = info.Dgv[visCols[i].Index, currentRowIndex].Value?.ToString() ?? "";
                            
                            RectangleF textRect = new RectangleF(rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4);
                            g.DrawString(val, fBody, Brushes.Black, textRect, fmtWrap);
                            
                            currX += colWidths[i];
                        }
                        y += maxRowHeight;
                        currentRowIndex++;
                    }
                    
                    y += 20; 
                    currentTableIndex++; 
                    currentRowIndex = 0;
                }
                e.HasMorePages = false; 
                currentTableIndex = 0; 
                currentRowIndex = 0;
            };

            PrintPreviewDialog ppd = new PrintPreviewDialog { Document = pd, Width = 1024, Height = 768, WindowState = FormWindowState.Maximized };
            ppd.ShowDialog();
        }
    }
}
