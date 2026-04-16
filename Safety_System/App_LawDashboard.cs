/// FILE: Safety_System/App_LawDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace Safety_System
{
    public class App_LawDashboard
    {
        private const string DbName = "法規";
        
        // 🟢 將「消防法規」納入系統掃描陣列中
        private readonly string[] _tableNames = { "環保法規", "職安衛法規", "消防法規", "其它法規" };
        
        private DataTable _dtAllLaws;
        private DataTable _dtDirectoryLaws; 
        private List<string> _errorLogs = new List<string>();
        
        private ComboBox _cboCategory;
        private ComboBox _cboYearlyCategory; 
        private ComboBox _cboYearlyYear;     
        private ComboBox _cboYearlyApplicability; 
        
        private DataGridView _dgvStats;
        private DataGridView _dgvCategoryLaws;
        private DataGridView _dgvThisYear;
        private Label _lblLoading; 

        public Control GetView()
        {
            TableLayoutPanel mainPanel = new TableLayoutPanel 
            { 
                Dock = DockStyle.Fill, 
                BackColor = Color.WhiteSmoke, 
                AutoScroll = true, 
                RowCount = 3, 
                ColumnCount = 1,
                Padding = new Padding(20) 
            };

            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 

            // ==========================================
            // 大框 1：年度法令總鑑別表
            // ==========================================
            GroupBox box1 = CreateDataBox("📌 年度法令總鑑別表查詢", 525); 
            
            Label lblTitle1 = new Label 
            { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n年度法令總鑑別表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DarkSlateBlue,
                Dock = DockStyle.Top,
                Height = 70
            };

            FlowLayoutPanel pnlAction1 = CreateActionFlowPanel("匯出 Excel", "匯出 PDF", 
                () => ExportToExcel(_dgvThisYear, "年度法令總鑑別表"), 
                () => ExportToPdf(_dgvThisYear, "年度法令總鑑別表", "年度法令總鑑別表"));
            
            Label lblCboCat1 = new Label { Text = "選擇類別:", AutoSize = true, Margin = new Padding(20, 8, 5, 0), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            _cboYearlyCategory = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 140, Margin = new Padding(0, 4, 0, 0) };
            _cboYearlyCategory.Items.AddRange(_tableNames);

            Label lblCboYear1 = new Label { Text = "查詢年度:", AutoSize = true, Margin = new Padding(15, 8, 5, 0), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            _cboYearlyYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 90, Margin = new Padding(0, 4, 0, 0) };
            
            int currentYear = DateTime.Now.Year;
            for (int i = 0; i <= 5; i++) {
                _cboYearlyYear.Items.Add((currentYear - i).ToString());
            }

            Label lblCboApp = new Label { Text = "適用性:", AutoSize = true, Margin = new Padding(15, 8, 5, 0), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            _cboYearlyApplicability = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 100, Margin = new Padding(0, 4, 0, 0) };
            _cboYearlyApplicability.Items.AddRange(new string[] { "全部", "適用", "不適用", "參考", "確認中", "" });

            _cboYearlyCategory.SelectedIndexChanged += (s, e) => { FilterYearlyLaws(); };
            _cboYearlyYear.SelectedIndexChanged += (s, e) => { FilterYearlyLaws(); };
            _cboYearlyApplicability.SelectedIndexChanged += (s, e) => { FilterYearlyLaws(); };

            pnlAction1.Controls.Add(lblCboCat1);
            pnlAction1.Controls.Add(_cboYearlyCategory);
            pnlAction1.Controls.Add(lblCboYear1);
            pnlAction1.Controls.Add(_cboYearlyYear);
            pnlAction1.Controls.Add(lblCboApp);
            pnlAction1.Controls.Add(_cboYearlyApplicability);

            _dgvThisYear = CreateStandardGrid();
            box1.Controls.Add(_dgvThisYear);
            box1.Controls.Add(pnlAction1);
            box1.Controls.Add(lblTitle1);
            _dgvThisYear.BringToFront(); 
            mainPanel.Controls.Add(box1, 0, 0);

            // ==========================================
            // 大框 2：統計摘要 (法令鑑別統計表)
            // ==========================================
            GroupBox box2 = CreateDataBox("📊 統計摘要", 400);
            Label lblTitle2 = new Label 
            { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n法令鑑別統計表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter, 
                ForeColor = Color.DarkSlateBlue,
                Dock = DockStyle.Top, 
                Height = 70 
            };

            FlowLayoutPanel pnlAction2 = CreateActionFlowPanel("匯出 Excel", "匯出 PDF", 
                () => ExportToExcel(_dgvStats, "統計摘要"), 
                () => ExportToPdf(_dgvStats, "統計摘要", "法令鑑別統計表"));
            
            _dgvStats = CreateStatsGrid();
            
            box2.Controls.Add(_dgvStats);
            box2.Controls.Add(pnlAction2);
            box2.Controls.Add(lblTitle2);
            _dgvStats.BringToFront();
            mainPanel.Controls.Add(box2, 0, 1);

            // ==========================================
            // 大框 3：目錄清單 (依類別檢視法令名稱)
            // ==========================================
            GroupBox box3 = CreateDataBox("📋 依類別檢視法令名稱", 450);

            Label lblTitle3 = new Label 
            { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n法令目錄一覽表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DarkSlateBlue,
                Dock = DockStyle.Top,
                Height = 70
            };

            FlowLayoutPanel pnlAction3 = CreateActionFlowPanel("匯出 Excel", "匯出 PDF", 
                () => ExportToExcel(_dgvCategoryLaws, "法令目錄"), 
                () => ExportToPdf(_dgvCategoryLaws, "法令目錄", "法令目錄一覽表"));
            
            pnlAction3.Controls.Add(new Label { Text = "選擇類別:", AutoSize = true, Margin = new Padding(20, 8, 5, 0), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) });
            _cboCategory = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 150, Margin = new Padding(0, 4, 0, 0) };
            _cboCategory.Items.AddRange(_tableNames);
            _cboCategory.SelectedIndexChanged += (s, e) => { FilterCategoryLaws(); };
            pnlAction3.Controls.Add(_cboCategory);

            Button btnSaveDir = new Button { Text = "💾 儲存確認日期", Size = new Size(200, 32), BackColor = Color.ForestGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(20, 2, 0, 0) };
            btnSaveDir.Click += BtnSaveDir_Click;
            pnlAction3.Controls.Add(btnSaveDir);

            _dgvCategoryLaws = CreateStandardGrid();
            _dgvCategoryLaws.ReadOnly = false; 
            
            _dgvCategoryLaws.CellClick += DgvCategoryLaws_CellClick;

            box3.Controls.Add(_dgvCategoryLaws);
            box3.Controls.Add(pnlAction3);
            box3.Controls.Add(lblTitle3);
            _dgvCategoryLaws.BringToFront();
            mainPanel.Controls.Add(box3, 0, 2);

            _lblLoading = new Label 
            {
                Text = "資料大量運算中，請稍候...",
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(200, 0, 0, 0),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            mainPanel.Controls.Add(_lblLoading);
            _lblLoading.BringToFront();

            LoadDashboardDataAsync();

            return mainPanel;
        }

        private async void LoadDashboardDataAsync()
        {
            DataTable dtStats = null;

            await Task.Run(() =>
            {
                try 
                {
                    _errorLogs.Clear();
                    LoadAndMergeData();
                    _dtDirectoryLaws = DataManager.GetTableData(DbName, "法規目錄一覽", "", "", "");
                    dtStats = BuildStatsData();
                } 
                catch (Exception ex) 
                {
                    _errorLogs.Add($"[背景運算異常] {ex.Message}");
                }
            });

            _lblLoading.Visible = false;

            if (dtStats != null) 
            {
                _dgvStats.DataSource = dtStats;
                FormatStatsGrid();
            }

            if (_cboYearlyCategory.Items.Count > 0) _cboYearlyCategory.SelectedIndex = 0;
            if (_cboYearlyYear.Items.Count > 0) _cboYearlyYear.SelectedIndex = 0;
            if (_cboYearlyApplicability.Items.Count > 0) _cboYearlyApplicability.SelectedIndex = 0; 
            
            FilterYearlyLaws(); 

            if (_cboCategory.Items.Count > 0) _cboCategory.SelectedIndex = 0;

            if (_errorLogs.Count > 0) 
            {
                MessageBox.Show("部分資料讀取異常。\n" + string.Join("\n", _errorLogs), "通知", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private GroupBox CreateDataBox(string title, int minHeight)
        {
            return new GroupBox 
            { 
                Text = title, 
                Dock = DockStyle.Fill, 
                MinimumSize = new Size(0, minHeight), 
                Margin = new Padding(0, 0, 0, 20), 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Padding = new Padding(15) 
            };
        }

        private FlowLayoutPanel CreateActionFlowPanel(string exText, string pdfText, Action exClick, Action pdfClick)
        {
            FlowLayoutPanel p = new FlowLayoutPanel 
            { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                MinimumSize = new Size(0, 45), 
                WrapContents = false, 
                Padding = new Padding(10, 5, 10, 5) 
            };
            
            Button btnEx = new Button { Text = "📊 " + exText, Size = new Size(130, 32), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 0, 10, 0) };
            Button btnPdf = new Button { Text = "📄 " + pdfText, Size = new Size(130, 32), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 0, 10, 0) };
            
            btnEx.Click += (s, e) => { exClick(); }; 
            btnPdf.Click += (s, e) => { pdfClick(); };
            
            p.Controls.Add(btnEx); 
            p.Controls.Add(btnPdf);
            return p;
        }

        private DataGridView CreateStandardGrid()
        {
            DataGridView dgv = new DataGridView 
            { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                AllowUserToDeleteRows = false, 
                ReadOnly = true, 
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F), 
                BorderStyle = BorderStyle.Fixed3D, 
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            };
            dgv.RowTemplate.Height = 35;
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            
            EnableDoubleBuffered(dgv); 
            return dgv;
        }

        private DataGridView CreateStatsGrid()
        {
            DataGridView dgv = CreateStandardGrid();
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.YellowGreen;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            dgv.DefaultCellStyle.BackColor = Color.White; 
            dgv.DefaultCellStyle.SelectionBackColor = Color.AliceBlue;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.GridColor = Color.Black;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; 
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dgv.ColumnHeadersHeight = 50; 
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

            return dgv;
        }

        private void EnableDoubleBuffered(DataGridView dgv)
        {
            typeof(DataGridView).InvokeMember("DoubleBuffered", 
                BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, 
                null, dgv, new object[] { true });
        }

        private void LoadAndMergeData()
        {
            _dtAllLaws = new DataTable();
            _dtAllLaws.Columns.Add("主分類", typeof(string));
            
            string[] expectedCols = { "Id", "日期", "法規名稱", "發布機關", "適用性", "鑑別日期", "有提升績效機會", "有潛在不符合風險" };
            foreach (var col in expectedCols) 
            {
                _dtAllLaws.Columns.Add(col, typeof(string));
            }

            foreach (string tbl in _tableNames) 
            {
                try 
                {
                    DataTable dt = DataManager.GetTableData(DbName, tbl, "", "", "");
                    if (dt == null || dt.Rows.Count == 0) continue; 
                    
                    foreach (DataRow row in dt.Rows) 
                    {
                        DataRow newRow = _dtAllLaws.NewRow();
                        newRow["主分類"] = tbl;
                        foreach (string col in expectedCols) 
                        {
                            if (dt.Columns.Contains(col) && row[col] != DBNull.Value) 
                            {
                                newRow[col] = row[col].ToString();
                            }
                            else 
                            {
                                newRow[col] = "";
                            }
                        }
                        _dtAllLaws.Rows.Add(newRow);
                    }
                } 
                catch { }
            }
        }

        private string GetSafeStr(DataRowView row, string colName) 
        { 
            if (row.Row.Table.Columns.Contains(colName) && row[colName] != DBNull.Value)
            {
                return row[colName].ToString().Trim();
            }
            return ""; 
        }

        private void FilterYearlyLaws()
        {
            if (_cboYearlyCategory.SelectedItem == null || 
                _cboYearlyYear.SelectedItem == null || 
                _cboYearlyApplicability.SelectedItem == null || 
                _dtAllLaws == null) return;
            
            string category = _cboYearlyCategory.SelectedItem.ToString();
            string year = _cboYearlyYear.SelectedItem.ToString();
            string applicability = _cboYearlyApplicability.SelectedItem.ToString();

            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("流水"); 
            dtShow.Columns.Add("法規名稱"); 
            dtShow.Columns.Add("日期");
            dtShow.Columns.Add("適用性");
            dtShow.Columns.Add("鑑別日期"); 

            DataView dv = new DataView(_dtAllLaws); 
            dv.RowFilter = $"主分類 = '{category}' AND 日期 LIKE '%{year}%'";

            var groupedData = new Dictionary<string, List<DataRowView>>();
            foreach (DataRowView drv in dv) 
            {
                string lawName = GetSafeStr(drv, "法規名稱");
                if (string.IsNullOrEmpty(lawName)) continue;
                if (!groupedData.ContainsKey(lawName)) groupedData[lawName] = new List<DataRowView>();
                groupedData[lawName].Add(drv);
            }

            int index = 1;
            foreach (var kvp in groupedData) 
            {
                string latestDate = "";
                string latestIdenDate = "";
                string firstApply = GetSafeStr(kvp.Value[0], "適用性");
                bool hasApplicable = false;
                
                foreach (var row in kvp.Value) 
                {
                    string d = GetSafeStr(row, "日期");
                    string iden = GetSafeStr(row, "鑑別日期");
                    string apply = GetSafeStr(row, "適用性");
                    
                    if (string.Compare(d, latestDate) > 0) latestDate = d;
                    if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                    if (apply == "適用") hasApplicable = true;
                }

                string finalApply = hasApplicable ? "適用" : firstApply;

                if (applicability != "全部" && finalApply != applicability) {
                    continue; 
                }
                
                dtShow.Rows.Add(index.ToString(), kvp.Key, latestDate, finalApply, latestIdenDate);
                index++;
            }

            _dgvThisYear.DataSource = dtShow;
            
            if (_dgvThisYear.Columns.Contains("流水")) {
                _dgvThisYear.Columns["流水"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvThisYear.Columns["流水"].Width = 70;
            }
            if (_dgvThisYear.Columns.Contains("法規名稱")) {
                _dgvThisYear.Columns["法規名稱"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            if (_dgvThisYear.Columns.Contains("日期")) {
                _dgvThisYear.Columns["日期"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvThisYear.Columns["日期"].Width = 140;
            }
            if (_dgvThisYear.Columns.Contains("適用性")) {
                _dgvThisYear.Columns["適用性"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvThisYear.Columns["適用性"].Width = 100;
            }
            if (_dgvThisYear.Columns.Contains("鑑別日期")) {
                _dgvThisYear.Columns["鑑別日期"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvThisYear.Columns["鑑別日期"].Width = 140;
            }
            
            _dgvThisYear.ClearSelection();
        }

        private DataTable BuildStatsData()
        {
            DataTable dtStats = new DataTable();
            string[] cols = { 
                "法規類別", "法規", "法條", "適用", 
                "參考", "不適用", "確認中", 
                "有提升績效機會", "有潛在不符合風險", "未鑑別" 
            };
            foreach (string c in cols) 
            {
                dtStats.Columns.Add(c);
            }

            int[] sums = new int[9];
            foreach (string cat in _tableNames) 
            {
                int[] v = new int[9];
                if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0) 
                {
                    DataView dv = new DataView(_dtAllLaws); 
                    dv.RowFilter = $"主分類 = '{cat}'";
                    
                    v[1] = dv.Count; 
                    HashSet<string> uniqueNames = new HashSet<string>();
                    
                    foreach (DataRowView row in dv) 
                    {
                        string name = GetSafeStr(row, "法規名稱");
                        string aStatus = GetSafeStr(row, "適用性");
                        string perfInc = GetSafeStr(row, "有提升績效機會");
                        string risk = GetSafeStr(row, "有潛在不符合風險");
                        
                        if (!string.IsNullOrEmpty(name)) uniqueNames.Add(name);
                        
                        if (aStatus == "適用") v[2]++; 
                        else if (aStatus == "參考") v[3]++; 
                        else if (aStatus == "不適用") v[4]++; 
                        else if (aStatus == "確認中") v[5]++; 
                        else if (string.IsNullOrEmpty(aStatus)) v[8]++;
                        
                        if (perfInc.ToLower() == "v") v[6]++; 
                        if (risk.ToLower() == "v") v[7]++;
                    }
                    v[0] = uniqueNames.Count; 
                }
                dtStats.Rows.Add(cat, v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8]);
                for (int i = 0; i < 9; i++) sums[i] += v[i];
            }
            
            dtStats.Rows.Add("合計", sums[0], sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], sums[7], sums[8]);
            return dtStats;
        }

        private void FormatStatsGrid()
        {
            if (_dgvStats.Rows.Count > 0) 
            {
                _dgvStats.Rows[_dgvStats.Rows.Count - 1].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold); 
                _dgvStats.Rows[_dgvStats.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Moccasin;
            }
            _dgvStats.ClearSelection();
        }

        private void FilterCategoryLaws()
        {
            if (_cboCategory.SelectedItem == null || _dtDirectoryLaws == null) return;
            string category = _cboCategory.SelectedItem.ToString();
            
            DataView dv = new DataView(_dtDirectoryLaws);
            dv.RowFilter = $"選項類別 = '{category}'";
            _dgvCategoryLaws.DataSource = dv.ToTable();
            
            if (_dgvCategoryLaws.Columns.Contains("Id")) _dgvCategoryLaws.Columns["Id"].Visible = false;
            if (_dgvCategoryLaws.Columns.Contains("選項類別")) _dgvCategoryLaws.Columns["選項類別"].Visible = false;

            foreach (DataGridViewColumn col in _dgvCategoryLaws.Columns) 
            {
                col.ReadOnly = (col.Name != "再次確認日期");
                if (col.Name == "再次確認日期") {
                    col.DefaultCellStyle.BackColor = Color.LightYellow; 
                }
            }

            if (_dgvCategoryLaws.Columns.Contains("流水號")) {
                _dgvCategoryLaws.Columns["流水號"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvCategoryLaws.Columns["流水號"].Width = 70;
            }
            if (_dgvCategoryLaws.Columns.Contains("法規名稱")) {
                _dgvCategoryLaws.Columns["法規名稱"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            if (_dgvCategoryLaws.Columns.Contains("日期")) {
                _dgvCategoryLaws.Columns["日期"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvCategoryLaws.Columns["日期"].Width = 140;
            }
            if (_dgvCategoryLaws.Columns.Contains("適用性")) {
                _dgvCategoryLaws.Columns["適用性"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvCategoryLaws.Columns["適用性"].Width = 100;
            }
            if (_dgvCategoryLaws.Columns.Contains("鑑別日期")) {
                _dgvCategoryLaws.Columns["鑑別日期"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvCategoryLaws.Columns["鑑別日期"].Width = 140;
            }
            if (_dgvCategoryLaws.Columns.Contains("再次確認日期")) {
                _dgvCategoryLaws.Columns["再次確認日期"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                _dgvCategoryLaws.Columns["再次確認日期"].Width = 170;
            }
            
            _dgvCategoryLaws.ClearSelection();
        }

        private void DgvCategoryLaws_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) 
            {
                if (_dgvCategoryLaws.Columns[e.ColumnIndex].Name == "再次確認日期") 
                {
                    var cell = _dgvCategoryLaws.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (cell.Value == DBNull.Value || string.IsNullOrWhiteSpace(cell.Value?.ToString())) 
                    {
                        cell.Value = DateTime.Today.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        private void BtnSaveDir_Click(object sender, EventArgs e)
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
            try {
                _dgvCategoryLaws.EndEdit();
                DataTable dt = (DataTable)_dgvCategoryLaws.DataSource;
                
                if (dt != null && DataManager.BulkSaveTable(DbName, "法規目錄一覽", dt)) {
                    MessageBox.Show("確認日期儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    _dtDirectoryLaws = DataManager.GetTableData(DbName, "法規目錄一覽", "", "", "");
                    FilterCategoryLaws(); 
                }
            } finally {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
            }
        }

        private void ExportToExcel(DataGridView dgv, string title)
        {
            if (dgv.Rows.Count == 0) 
            { 
                MessageBox.Show("沒有資料可匯出！"); 
                return; 
            }
            
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = title + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = new DataTable();
                        List<DataGridViewColumn> visCols = new List<DataGridViewColumn>();

                        foreach (DataGridViewColumn col in dgv.Columns) 
                        {
                            if (col.Visible) 
                            {
                                visCols.Add(col);
                                dt.Columns.Add(col.HeaderText.Replace("\n", ""));
                            }
                        }
                        
                        foreach (DataGridViewRow row in dgv.Rows) 
                        {
                            if (row.IsNewRow) continue;
                            DataRow dRow = dt.NewRow();
                            for (int i = 0; i < visCols.Count; i++) 
                            {
                                var cellVal = row.Cells[visCols[i].Index].Value;
                                dRow[i] = cellVal != null ? cellVal.ToString() : "";
                            }
                            dt.Rows.Add(dRow);
                        }

                        using (ExcelPackage p = new ExcelPackage()) 
                        {
                            var ws = p.Workbook.Worksheets.Add("Data");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns(); 
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("Excel 匯出成功！");
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("Excel 匯出失敗：" + ex.Message); 
                    }
                }
            }
        }

        private void ExportToPdf(DataGridView dgv, string fileName, string reportTitle)
        {
            if (dgv.Rows.Count == 0) 
            { 
                MessageBox.Show("沒有資料可列印！"); 
                return; 
            }
            
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = fileName + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = true; 
                    pd.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                    int rowIndex = 0;
                    pd.PrintPage += (s, e) => 
                    {
                        Graphics g = e.Graphics;
                        Font font = new Font("Microsoft JhengHei UI", 9F);
                        Font headerFont = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
                        Font titleFont = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                        Font dateFont = new Font("Microsoft JhengHei UI", 11F);
                        
                        StringFormat fmtCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                        StringFormat fmtLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };

                        float y = e.MarginBounds.Top;
                        float totalWidth = 0;
                        
                        List<DataGridViewColumn> visCols = new List<DataGridViewColumn>();
                        foreach (DataGridViewColumn col in dgv.Columns) 
                        {
                            if (col.Visible) 
                            {
                                visCols.Add(col);
                                totalWidth += col.Width;
                            }
                        }

                        float scale = e.MarginBounds.Width / totalWidth;
                        if (scale > 1f) scale = 1f; 

                        g.ScaleTransform(scale, scale);
                        float scaledHeight = e.MarginBounds.Height / scale;
                        float scaledWidth = e.MarginBounds.Width / scale;
                        float x = e.MarginBounds.Left / scale;

                        string companyTitle = "台灣玻璃工業股份有限公司-彰濱廠\n" + reportTitle;
                        string exportDate = "導出日期: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                        SizeF titleSize = g.MeasureString(companyTitle, titleFont, (int)scaledWidth, fmtCenter);
                        RectangleF titleRect = new RectangleF(x, y / scale, scaledWidth, titleSize.Height + 10);
                        g.DrawString(companyTitle, titleFont, Brushes.DarkSlateBlue, titleRect, fmtCenter);
                        y += (titleSize.Height + 10) * scale;

                        SizeF dateSize = g.MeasureString(exportDate, dateFont, (int)scaledWidth, fmtLeft);
                        RectangleF dateRect = new RectangleF(x, y / scale, scaledWidth, dateSize.Height + 10);
                        g.DrawString(exportDate, dateFont, Brushes.Black, dateRect, fmtLeft);
                        y += (dateSize.Height + 15) * scale;

                        float headerH = dgv.ColumnHeadersHeight < 40 ? 40 : dgv.ColumnHeadersHeight;
                        
                        for (int i = 0; i < visCols.Count; i++) 
                        {
                            RectangleF rectF = new RectangleF(x, y / scale, visCols[i].Width, headerH);
                            Rectangle rect = Rectangle.Round(rectF); 
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, rect);
                            
                            string headerText = visCols[i].HeaderText.Replace("\n", "");
                            g.DrawString(headerText, headerFont, Brushes.Black, rect, fmtCenter);
                            x += visCols[i].Width;
                        }
                        y += headerH * scale;

                        while (rowIndex < dgv.Rows.Count) 
                        {
                            DataGridViewRow row = dgv.Rows[rowIndex];
                            float rowH = row.Height < 30 ? 30 : row.Height;
                            
                            if ((y / scale) + rowH > scaledHeight + (e.MarginBounds.Top / scale)) 
                            {
                                e.HasMorePages = true; 
                                return;
                            }

                            x = e.MarginBounds.Left / scale;
                            for (int i = 0; i < visCols.Count; i++) 
                            {
                                RectangleF rectF = new RectangleF(x, y / scale, visCols[i].Width, rowH);
                                Rectangle rect = Rectangle.Round(rectF);
                                g.DrawRectangle(Pens.Black, rect);
                                
                                string val = row.Cells[visCols[i].Index].Value?.ToString() ?? "";
                                g.DrawString(val, font, Brushes.Black, rect, fmtCenter);
                                x += visCols[i].Width;
                            }
                            y += rowH * scale;
                            rowIndex++;
                        }
                        e.HasMorePages = false;
                        rowIndex = 0; 
                    };

                    try 
                    { 
                        pd.Print(); 
                        MessageBox.Show("PDF 匯出成功！"); 
                    }
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message); 
                    }
                }
            }
        }
    }
}
