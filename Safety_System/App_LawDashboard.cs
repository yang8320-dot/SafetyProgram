/// FILE: Safety_System/App_LawDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_LawDashboard
    {
        private const string DbName = "法規";
        private readonly string[] _tableNames = { "環保法規", "職安衛法規", "其它法規" };
        
        // 記憶體中的快取資料 
        private DataTable _dtAllLaws;
        
        // 錯誤訊息收集器
        private List<string> _errorLogs = new List<string>();
        
        // 第三框的 UI 控制項
        private ComboBox _cboCategory;
        private DataGridView _dgvCategoryLaws;

        public Control GetView()
        {
            _errorLogs.Clear();

            // 1. 安全載入資料庫
            try { LoadAndMergeData(); }
            catch (Exception ex) { _errorLogs.Add($"[資料讀取階段] {ex.Message}"); }

            // 🟢 改用 TableLayoutPanel 徹底鎖定 1, 2, 3 區塊的上下順序
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
            // 第一區塊：今年修正法規 (最上方)
            // ==========================================
            GroupBox box1 = CreateGroupBox("📌 今年修正法規一覽 (排除重複名稱，依適用性權重顯示)", 300);
            DataGridView dgvThisYear = CreateStandardGrid();
            try { PopulateThisYearData(dgvThisYear); }
            catch (Exception ex) { _errorLogs.Add($"[第一框-今年修正法規] 發生異常：{ex.Message}"); }
            box1.Controls.Add(dgvThisYear);
            
            mainPanel.Controls.Add(box1, 0, 0); // 強制放在第 0 列

            // ==========================================
            // 第二區塊：統計摘要 (中間)
            // ==========================================
            GroupBox box2 = CreateGroupBox("📊 環安衛法令及其他要求內容一覽表 (統計摘要)", 350);
            box2.Padding = new Padding(15, 40, 15, 15);
            
            Label lblTitle2 = new Label { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n環安衛法令及其他要求內容一覽表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter, 
                Dock = DockStyle.Top, 
                Height = 60 
            };
            
            DataGridView dgvStats = CreateStatsGrid();
            try { PopulateStatsData(dgvStats); }
            catch (Exception ex) { _errorLogs.Add($"[第二框-統計摘要] 發生異常：{ex.Message}"); }
            
            // 🟢 修正遮擋問題：先加 Label，再加 Grid，並把 Grid 移到最上層以填滿剩餘空間
            box2.Controls.Add(lblTitle2);
            box2.Controls.Add(dgvStats);
            dgvStats.BringToFront(); 

            mainPanel.Controls.Add(box2, 0, 1); // 強制放在第 1 列

            // ==========================================
            // 第三區塊：依類別清單 (最下方)
            // ==========================================
            GroupBox box3 = CreateGroupBox("📋 依類別檢視法令名稱一覽", 400);
            box3.Padding = new Padding(15, 40, 15, 15);

            Panel pnlTop3 = new Panel { Dock = DockStyle.Top, Height = 60 };
            
            // 🟢 移除紅字，調整下拉選單位置
            _cboCategory = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 200, Location = new Point(10, 15) };
            _cboCategory.Items.AddRange(_tableNames);
            
            _cboCategory.SelectedIndexChanged += (s, e) => {
                try { PopulateCategoryLaws(); } 
                catch (Exception ex) { MessageBox.Show($"查詢目錄時發生錯誤：{ex.Message}", "查詢異常", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };

            Label lblTitle3 = new Label { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n環安衛法令及其他要求目錄一覽表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill // 🟢 自動置中填滿
            };

            pnlTop3.Controls.Add(_cboCategory);
            pnlTop3.Controls.Add(lblTitle3);
            _cboCategory.BringToFront(); // 確保選單在標題上方可被點擊

            _dgvCategoryLaws = CreateStandardGrid();
            
            // 🟢 修正遮擋問題
            box3.Controls.Add(pnlTop3);
            box3.Controls.Add(_dgvCategoryLaws);
            _dgvCategoryLaws.BringToFront();

            mainPanel.Controls.Add(box3, 0, 2); // 強制放在第 2 列

            // 初始載入第三框資料
            try { if (_cboCategory.Items.Count > 0) _cboCategory.SelectedIndex = 0; }
            catch (Exception ex) { _errorLogs.Add($"[第三框-目錄一覽表] 發生異常：{ex.Message}"); }

            // 若有錯誤，跳出提示視窗
            if (_errorLogs.Count > 0)
            {
                string msg = "看板載入時有部分資料讀取異常。\n(可能原因：資料表尚未建立、缺少必填欄位或資料為空)\n\n系統已略過錯誤項目，保留正常顯示的部分。\n\n詳細錯誤如下：\n" 
                             + string.Join("\n", _errorLogs);
                MessageBox.Show(msg, "看板載入異常通知", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return mainPanel;
        }

        // ==========================================
        // 資料讀取與合併 (防呆)
        // ==========================================
        private void LoadAndMergeData()
        {
            _dtAllLaws = new DataTable();
            _dtAllLaws.Columns.Add("主分類", typeof(string));
            
            string[] expectedCols = { "Id", "日期", "法規名稱", "發布機關", "施行日期", "合規狀態", "適用性", "鑑別日期" };
            foreach (var col in expectedCols) _dtAllLaws.Columns.Add(col, typeof(string));

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
                                newRow[col] = row[col].ToString();
                            else 
                                newRow[col] = ""; 
                        }
                        _dtAllLaws.Rows.Add(newRow);
                    }
                }
                catch (Exception ex)
                {
                    _errorLogs.Add($"- 無法讀取表【{tbl}】 ({ex.Message})");
                }
            }
        }

        private string GetSafeStr(DataRowView row, string colName)
        {
            if (row.Row.Table.Columns.Contains(colName) && row[colName] != DBNull.Value)
                return row[colName].ToString().Trim();
            return "";
        }

        // ==========================================
        // 第一區塊：今年修正法規
        // ==========================================
        private void PopulateThisYearData(DataGridView dgv)
        {
            try 
            {
                DataTable dtShow = new DataTable();
                dtShow.Columns.Add("日期");
                dtShow.Columns.Add("鑑別日期");
                dtShow.Columns.Add("法規名稱");
                dtShow.Columns.Add("適用性");

                if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0)
                {
                    string currentYear = DateTime.Now.Year.ToString();
                    DataView dv = new DataView(_dtAllLaws);
                    dv.RowFilter = $"日期 LIKE '%{currentYear}%'";

                    Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
                    
                    foreach (DataRowView drv in dv)
                    {
                        string lawName = GetSafeStr(drv, "法規名稱");
                        if (string.IsNullOrEmpty(lawName)) continue;
                        
                        if (!groupedData.ContainsKey(lawName))
                            groupedData[lawName] = new List<DataRowView>();
                        
                        groupedData[lawName].Add(drv);
                    }

                    foreach (var kvp in groupedData)
                    {
                        string lawName = kvp.Key;
                        var list = kvp.Value;
                        if (list.Count == 0) continue; 
                        
                        string latestDate = "";
                        string latestIdenDate = "";
                        bool hasApplicable = false;
                        string firstApply = GetSafeStr(list[0], "適用性");

                        foreach (var row in list)
                        {
                            string d = GetSafeStr(row, "日期");
                            string iden = GetSafeStr(row, "鑑別日期");
                            string apply = GetSafeStr(row, "適用性");

                            if (string.Compare(d, latestDate) > 0) latestDate = d;
                            if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                            if (apply == "適用") hasApplicable = true;
                        }

                        string finalApply = hasApplicable ? "適用" : firstApply;
                        dtShow.Rows.Add(latestDate, latestIdenDate, lawName, finalApply);
                    }
                }

                dgv.DataSource = dtShow;
                if (dgv.Columns.Count >= 4) dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch { /* 安全略過 */ }
        }

        // ==========================================
        // 第二區塊：統計摘要
        // ==========================================
        private void PopulateStatsData(DataGridView dgv)
        {
            try
            {
                DataTable dtStats = new DataTable();
                dtStats.Columns.Add("類別");
                dtStats.Columns.Add("收集【法規】數");
                dtStats.Columns.Add("收集【法條】數");
                dtStats.Columns.Add("【適用】法條數");
                dtStats.Columns.Add("【參考】數");
                dtStats.Columns.Add("【不適用】法條數");
                dtStats.Columns.Add("【確認中】數");
                dtStats.Columns.Add("合法且有提升\n績效機會法條數");
                dtStats.Columns.Add("合法但潛在不\n符合風險法條數");
                dtStats.Columns.Add("【未鑑別】\n法條數");

                int sumLaws = 0, sumItems = 0, sumApply = 0, sumRef = 0, sumNotApply = 0, sumCheck = 0, sumGood = 0, sumRisk = 0, sumUnk = 0;

                foreach (string cat in _tableNames)
                {
                    int uniqueLaws = 0, items = 0, apply = 0, refer = 0, notApply = 0, checking = 0, good = 0, risk = 0, unk = 0;

                    if (_dtAllLaws != null && _dtAllLaws.Rows.Count > 0)
                    {
                        DataView dv = new DataView(_dtAllLaws);
                        dv.RowFilter = $"主分類 = '{cat}'";

                        items = dv.Count;
                        HashSet<string> uniqueNames = new HashSet<string>();

                        foreach (DataRowView row in dv)
                        {
                            string name = GetSafeStr(row, "法規名稱");
                            if (!string.IsNullOrEmpty(name)) uniqueNames.Add(name);

                            string aStatus = GetSafeStr(row, "適用性");
                            string cStatus = GetSafeStr(row, "合規狀態");

                            if (aStatus == "適用") apply++;
                            else if (aStatus == "參考") refer++;
                            else if (aStatus == "不適用") notApply++;
                            else if (aStatus == "確認中") checking++;
                            else if (string.IsNullOrEmpty(aStatus)) unk++;

                            if (cStatus.Contains("提升")) good++;
                            if (cStatus.Contains("潛在不符合")) risk++;
                        }
                        uniqueLaws = uniqueNames.Count;
                    }

                    dtStats.Rows.Add(cat, uniqueLaws, items, apply, refer, notApply, checking, good, risk, unk);

                    sumLaws += uniqueLaws; sumItems += items; sumApply += apply; sumRef += refer; sumNotApply += notApply; 
                    sumCheck += checking; sumGood += good; sumRisk += risk; sumUnk += unk;
                }

                dtStats.Rows.Add("合計", sumLaws, sumItems, sumApply, sumRef, sumNotApply, sumCheck, sumGood, sumRisk, sumUnk);

                dgv.DataSource = dtStats;
                
                if (dgv.Columns.Count > 0) dgv.Columns[0].HeaderText = "環保法規\n(類別)";

                dgv.DataBindingComplete += (s, e) => {
                    if (dgv.Rows.Count > 0) {
                        int lastIndex = dgv.Rows.Count - 1;
                        if (lastIndex >= 0) dgv.Rows[lastIndex].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                    }
                };
            }
            catch { /* 安全略過 */ }
        }

        // ==========================================
        // 第三區塊：依類別檢視
        // ==========================================
        private void PopulateCategoryLaws()
        {
            try
            {
                DataTable dtShow = new DataTable();
                dtShow.Columns.Add("流水號");
                dtShow.Columns.Add("法令名稱");
                dtShow.Columns.Add("公告日");
                dtShow.Columns.Add("鑑別日期");

                if (_cboCategory.SelectedItem != null && _dtAllLaws != null && _dtAllLaws.Rows.Count > 0)
                {
                    string selectedCat = _cboCategory.SelectedItem.ToString();

                    DataView dv = new DataView(_dtAllLaws);
                    dv.RowFilter = $"主分類 = '{selectedCat}'";

                    Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
                    
                    foreach (DataRowView drv in dv)
                    {
                        string lawName = GetSafeStr(drv, "法規名稱");
                        if (string.IsNullOrEmpty(lawName)) continue;
                        
                        if (!groupedData.ContainsKey(lawName))
                            groupedData[lawName] = new List<DataRowView>();
                        
                        groupedData[lawName].Add(drv);
                    }

                    int index = 1;
                    foreach (var kvp in groupedData)
                    {
                        string lawName = kvp.Key;
                        var list = kvp.Value;
                        if (list.Count == 0) continue;
                        
                        string latestDate = "";
                        string latestIdenDate = "";

                        foreach (var row in list)
                        {
                            string d = GetSafeStr(row, "日期");
                            string iden = GetSafeStr(row, "鑑別日期");

                            if (string.Compare(d, latestDate) > 0) latestDate = d;
                            if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                        }

                        dtShow.Rows.Add(index.ToString(), lawName, latestDate, latestIdenDate);
                        index++;
                    }
                }

                _dgvCategoryLaws.DataSource = dtShow;
                
                if (_dgvCategoryLaws.Columns.Count >= 4)
                {
                    _dgvCategoryLaws.Columns[0].Width = 80;
                    _dgvCategoryLaws.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    _dgvCategoryLaws.Columns[2].Width = 150;
                    _dgvCategoryLaws.Columns[3].Width = 150;
                }
            }
            catch { /* 安全略過 */ }
        }

        // ==========================================
        // UI 建立輔助方法
        // ==========================================
        private GroupBox CreateGroupBox(string title, int minHeight)
        {
            return new GroupBox 
            { 
                Text = title, 
                Dock = DockStyle.Fill, 
                MinimumSize = new Size(0, minHeight), 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Padding = new Padding(15),
                Margin = new Padding(0, 0, 0, 20)
            };
        }

        private DataGridView CreateStandardGrid()
        {
            return new DataGridView 
            { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                Font = new Font("Microsoft JhengHei UI", 11F),
                BorderStyle = BorderStyle.Fixed3D
            };
        }

        private DataGridView CreateStatsGrid()
        {
            DataGridView dgv = CreateStandardGrid();
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.YellowGreen;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

            dgv.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.SelectionBackColor = Color.Khaki;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.GridColor = Color.Black;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            return dgv;
        }
    }
}
