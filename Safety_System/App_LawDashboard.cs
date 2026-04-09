/// FILE: Safety_System/App_LawDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_LawDashboard
    {
        private const string DbName = "法規";
        private readonly string[] _tableNames = { "環保法規", "職安衛法規", "其它法規" };
        
        // 記憶體中的快取資料 (將三個表合併成一個大表)
        private DataTable _dtAllLaws;
        
        // 第三框的 UI 控制項
        private ComboBox _cboCategory;
        private DataGridView _dgvCategoryLaws;

        public Control GetView()
        {
            LoadAndMergeData();

            // 主版面配置
            Panel mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };

            // ==========================================
            // 第一區塊：今年修正法規
            // ==========================================
            GroupBox box1 = CreateGroupBox("📌 今年修正法規一覽 (排除重複名稱，依適用性權重顯示)");
            DataGridView dgvThisYear = CreateStandardGrid();
            PopulateThisYearData(dgvThisYear);
            box1.Controls.Add(dgvThisYear);
            mainPanel.Controls.Add(box1);

            // ==========================================
            // 第二區塊：統計摘要 (紅框)
            // ==========================================
            GroupBox box2 = CreateGroupBox("📊 環安衛法令及其他要求內容一覽表 (統計摘要)");
            box2.Padding = new Padding(15, 30, 15, 15);
            
            Label lblTitle2 = new Label { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n環安衛法令及其他要求內容一覽表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter, 
                Dock = DockStyle.Top, 
                Height = 60 
            };
            
            DataGridView dgvStats = CreateStatsGrid();
            PopulateStatsData(dgvStats);
            
            box2.Controls.Add(dgvStats);
            box2.Controls.Add(lblTitle2);
            mainPanel.Controls.Add(box2);

            // ==========================================
            // 第三區塊：依類別清單 (藍框)
            // ==========================================
            GroupBox box3 = CreateGroupBox("📋 依類別檢視法令名稱一覽");
            box3.Padding = new Padding(15, 30, 15, 15);

            Panel pnlTop3 = new Panel { Dock = DockStyle.Top, Height = 100 };
            
            Label lblCboTitle = new Label { Text = "下拉選單 (環保法規、職安衛法規、其它法規)：", ForeColor = Color.Red, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Location = new Point(10, 10) };
            _cboCategory = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 200, Location = new Point(10, 40) };
            _cboCategory.Items.AddRange(_tableNames);
            _cboCategory.SelectedIndex = 0;
            _cboCategory.SelectedIndexChanged += (s, e) => PopulateCategoryLaws();

            Label lblTitle3 = new Label { 
                Text = "台灣玻璃工業股份有限公司-彰濱廠\n法令名稱一覽表", 
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), 
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = true,
                Location = new Point(400, 30)
            };

            pnlTop3.Controls.Add(lblCboTitle);
            pnlTop3.Controls.Add(_cboCategory);
            pnlTop3.Controls.Add(lblTitle3);

            _dgvCategoryLaws = CreateStandardGrid();
            PopulateCategoryLaws(); // 初始載入

            box3.Controls.Add(_dgvCategoryLaws);
            box3.Controls.Add(pnlTop3);
            mainPanel.Controls.Add(box3);

            box3.BringToFront();
            box2.BringToFront();
            box1.BringToFront();

            return mainPanel;
        }

        // ==========================================
        // 資料讀取與合併
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
                    if (dt == null) continue;

                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow newRow = _dtAllLaws.NewRow();
                        newRow["主分類"] = tbl;
                        foreach (string col in expectedCols)
                        {
                            if (dt.Columns.Contains(col)) newRow[col] = row[col]?.ToString() ?? "";
                            else newRow[col] = "";
                        }
                        _dtAllLaws.Rows.Add(newRow);
                    }
                }
                catch { /* 忽略尚未建立的表 */ }
            }
        }

        // ==========================================
        // 修正版：第一區塊 (避免使用 AsEnumerable)
        // ==========================================
        private void PopulateThisYearData(DataGridView dgv)
        {
            string currentYear = DateTime.Now.Year.ToString();
            
            // 使用 DataView 篩選今年度資料
            DataView dv = new DataView(_dtAllLaws);
            dv.RowFilter = $"日期 LIKE '%{currentYear}%'";

            // 手動分群 (法規名稱 -> 存放多筆資料的清單)
            Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
            
            foreach (DataRowView drv in dv)
            {
                string lawName = drv["法規名稱"].ToString().Trim();
                if (string.IsNullOrEmpty(lawName)) continue;
                
                if (!groupedData.ContainsKey(lawName))
                    groupedData[lawName] = new List<DataRowView>();
                
                groupedData[lawName].Add(drv);
            }

            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("日期");
            dtShow.Columns.Add("鑑別日期");
            dtShow.Columns.Add("法規名稱");
            dtShow.Columns.Add("適用性");

            // 分析每一組法規並排序取得最新資料
            foreach (var kvp in groupedData)
            {
                string lawName = kvp.Key;
                var list = kvp.Value;
                
                // 找出該群組中最新的日期與鑑別日期 (字串降冪排序)
                string latestDate = "";
                string latestIdenDate = "";
                bool hasApplicable = false;
                string firstApply = list[0]["適用性"].ToString();

                foreach (var row in list)
                {
                    string d = row["日期"].ToString();
                    string iden = row["鑑別日期"].ToString();
                    string apply = row["適用性"].ToString();

                    if (string.Compare(d, latestDate) > 0) latestDate = d;
                    if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                    if (apply == "適用") hasApplicable = true;
                }

                // 若群組中任一筆為「適用」，即為「適用」，否則顯示第一筆的狀態
                string finalApply = hasApplicable ? "適用" : firstApply;

                dtShow.Rows.Add(latestDate, latestIdenDate, lawName, finalApply);
            }

            dgv.DataSource = dtShow;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        // ==========================================
        // 修正版：第二區塊 (統計摘要)
        // ==========================================
        private void PopulateStatsData(DataGridView dgv)
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
                DataView dv = new DataView(_dtAllLaws);
                dv.RowFilter = $"主分類 = '{cat}'";

                int items = dv.Count;
                HashSet<string> uniqueNames = new HashSet<string>();
                int apply = 0, refer = 0, notApply = 0, checking = 0, good = 0, risk = 0, unk = 0;

                foreach (DataRowView row in dv)
                {
                    string name = row["法規名稱"].ToString().Trim();
                    if (!string.IsNullOrEmpty(name)) uniqueNames.Add(name);

                    string aStatus = row["適用性"].ToString();
                    string cStatus = row["合規狀態"].ToString();

                    if (aStatus == "適用") apply++;
                    else if (aStatus == "參考") refer++;
                    else if (aStatus == "不適用") notApply++;
                    else if (aStatus == "確認中") checking++;
                    else if (string.IsNullOrEmpty(aStatus)) unk++;

                    if (cStatus.Contains("提升")) good++;
                    if (cStatus.Contains("潛在不符合")) risk++;
                }

                int uniqueLaws = uniqueNames.Count;
                dtStats.Rows.Add(cat, uniqueLaws, items, apply, refer, notApply, checking, good, risk, unk);

                // 加總
                sumLaws += uniqueLaws; sumItems += items; sumApply += apply; sumRef += refer; sumNotApply += notApply; 
                sumCheck += checking; sumGood += good; sumRisk += risk; sumUnk += unk;
            }

            // 加入合計列
            dtStats.Rows.Add("合計", sumLaws, sumItems, sumApply, sumRef, sumNotApply, sumCheck, sumGood, sumRisk, sumUnk);

            dgv.DataSource = dtStats;
            dgv.Columns[0].HeaderText = "環保法規\n(類別)";

            dgv.DataBindingComplete += (s, e) => {
                if (dgv.Rows.Count > 0) {
                    dgv.Rows[dgv.Rows.Count - 1].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                }
            };
        }

        // ==========================================
        // 修正版：第三區塊 (依類別檢視)
        // ==========================================
        private void PopulateCategoryLaws()
        {
            if (_cboCategory.SelectedItem == null) return;
            string selectedCat = _cboCategory.SelectedItem.ToString();

            DataView dv = new DataView(_dtAllLaws);
            dv.RowFilter = $"主分類 = '{selectedCat}'";

            Dictionary<string, List<DataRowView>> groupedData = new Dictionary<string, List<DataRowView>>();
            
            foreach (DataRowView drv in dv)
            {
                string lawName = drv["法規名稱"].ToString().Trim();
                if (string.IsNullOrEmpty(lawName)) continue;
                
                if (!groupedData.ContainsKey(lawName))
                    groupedData[lawName] = new List<DataRowView>();
                
                groupedData[lawName].Add(drv);
            }

            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("流水號");
            dtShow.Columns.Add("法令名稱");
            dtShow.Columns.Add("公告日");
            dtShow.Columns.Add("鑑別日期");

            int index = 1;
            foreach (var kvp in groupedData)
            {
                string lawName = kvp.Key;
                var list = kvp.Value;
                
                string latestDate = "";
                string latestIdenDate = "";

                foreach (var row in list)
                {
                    string d = row["日期"].ToString();
                    string iden = row["鑑別日期"].ToString();

                    if (string.Compare(d, latestDate) > 0) latestDate = d;
                    if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                }

                dtShow.Rows.Add(index.ToString(), lawName, latestDate, latestIdenDate);
                index++;
            }

            _dgvCategoryLaws.DataSource = dtShow;
            
            if (_dgvCategoryLaws.Columns.Count > 0)
            {
                _dgvCategoryLaws.Columns[0].Width = 80;
                _dgvCategoryLaws.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                _dgvCategoryLaws.Columns[2].Width = 150;
                _dgvCategoryLaws.Columns[3].Width = 150;
            }
        }

        // ==========================================
        // UI 建立輔助方法
        // ==========================================
        private GroupBox CreateGroupBox(string title)
        {
            return new GroupBox 
            { 
                Text = title, 
                Dock = DockStyle.Top, 
                Height = 300, 
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
