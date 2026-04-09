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
        
        // 記憶體中的快取資料 (將三個表合併成一個大表方便運算)
        private DataTable _dtAllLaws;
        
        // 第三框的 UI 控制項
        private ComboBox _cboCategory;
        private DataGridView _dgvCategoryLaws;

        public Control GetView()
        {
            // 1. 載入並合併所有法規資料
            LoadAndMergeData();

            // 2. 主版面配置 (可滾動的 Panel)
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
            // 第二區塊：環安衛法令及其他要求內容一覽表 (紅框)
            // ==========================================
            GroupBox box2 = CreateGroupBox("📊 環安衛法令及其他要求內容一覽表 (統計摘要)");
            box2.Padding = new Padding(15, 30, 15, 15);
            
            // 標題 Label
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
            // 第三區塊：法令名稱一覽表 (藍框)
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
                Location = new Point(400, 30) // 粗略置中
            };

            pnlTop3.Controls.Add(lblCboTitle);
            pnlTop3.Controls.Add(_cboCategory);
            pnlTop3.Controls.Add(lblTitle3);

            _dgvCategoryLaws = CreateStandardGrid();
            PopulateCategoryLaws(); // 初始載入

            box3.Controls.Add(_dgvCategoryLaws);
            box3.Controls.Add(pnlTop3);
            mainPanel.Controls.Add(box3);

            // 確保順序是由上到下
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
            
            // 定義預期欄位，確保 DataTable 結構一致
            string[] expectedCols = { "Id", "日期", "法規名稱", "發布機關", "施行日期", "適用性", "合規狀態", "鑑別日期" };
            foreach (var col in expectedCols) _dtAllLaws.Columns.Add(col, typeof(string));

            foreach (string tbl in _tableNames)
            {
                try
                {
                    // 讀取該表所有資料
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
                catch { /* 忽略尚未建立的資料表 */ }
            }
        }

        // ==========================================
        // 第一區塊邏輯：今年修正法規
        // ==========================================
        private void PopulateThisYearData(DataGridView dgv)
        {
            string currentYear = DateTime.Now.Year.ToString();

            // 1. 篩選出今年度的資料
            var thisYearRows = _dtAllLaws.AsEnumerable()
                .Where(r => r.Field<string>("日期").Contains(currentYear));

            // 2. 依照「法規名稱」群組，排除重複
            var groupedData = thisYearRows.GroupBy(r => r.Field<string>("法規名稱").Trim())
                .Where(g => !string.IsNullOrEmpty(g.Key))
                .Select(g => new
                {
                    法規名稱 = g.Key,
                    日期 = g.OrderByDescending(r => r.Field<string>("日期")).First().Field<string>("日期"),
                    鑑別日期 = g.OrderByDescending(r => r.Field<string>("鑑別日期")).First().Field<string>("鑑別日期"),
                    // 權重判斷：只要有一筆是「適用」，結果就是「適用」
                    適用性 = g.Any(r => r.Field<string>("適用性") == "適用") ? "適用" : g.First().Field<string>("適用性")
                }).ToList();

            // 3. 綁定到 Grid
            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("日期");
            dtShow.Columns.Add("鑑別日期");
            dtShow.Columns.Add("法規名稱");
            dtShow.Columns.Add("適用性");

            foreach (var item in groupedData)
            {
                dtShow.Rows.Add(item.日期, item.鑑別日期, item.法規名稱, item.適用性);
            }

            dgv.DataSource = dtShow;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        // ==========================================
        // 第二區塊邏輯：統計摘要 (紅框)
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
                var rows = _dtAllLaws.AsEnumerable().Where(r => r.Field<string>("主分類") == cat).ToList();
                
                int uniqueLaws = rows.Select(r => r.Field<string>("法規名稱").Trim()).Where(n => !string.IsNullOrEmpty(n)).Distinct().Count();
                int items = rows.Count;
                int apply = rows.Count(r => r.Field<string>("適用性") == "適用");
                int refer = rows.Count(r => r.Field<string>("適用性") == "參考");
                int notApply = rows.Count(r => r.Field<string>("適用性") == "不適用");
                int checking = rows.Count(r => r.Field<string>("適用性") == "確認中");
                int good = rows.Count(r => r.Field<string>("合規狀態").Contains("提升"));
                int risk = rows.Count(r => r.Field<string>("合規狀態").Contains("潛在不符合"));
                int unk = rows.Count(r => string.IsNullOrEmpty(r.Field<string>("適用性")));

                dtStats.Rows.Add(cat, uniqueLaws, items, apply, refer, notApply, checking, good, risk, unk);

                // 加總
                sumLaws += uniqueLaws; sumItems += items; sumApply += apply; sumRef += refer; sumNotApply += notApply; 
                sumCheck += checking; sumGood += good; sumRisk += risk; sumUnk += unk;
            }

            // 加入合計列
            dtStats.Rows.Add("合計", sumLaws, sumItems, sumApply, sumRef, sumNotApply, sumCheck, sumGood, sumRisk, sumUnk);

            dgv.DataSource = dtStats;
            
            // 替換第一欄的標題以符合圖片
            dgv.Columns[0].HeaderText = "環保法規\n(類別)";

            // 針對合計列進行粗體顯示
            dgv.DataBindingComplete += (s, e) => {
                if (dgv.Rows.Count > 0) {
                    dgv.Rows[dgv.Rows.Count - 1].DefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold);
                }
            };
        }

        // ==========================================
        // 第三區塊邏輯：依類別清單 (藍框)
        // ==========================================
        private void PopulateCategoryLaws()
        {
            if (_cboCategory.SelectedItem == null) return;
            string selectedCat = _cboCategory.SelectedItem.ToString();

            var rows = _dtAllLaws.AsEnumerable()
                .Where(r => r.Field<string>("主分類") == selectedCat);

            // 群組去重複
            var groupedData = rows.GroupBy(r => r.Field<string>("法規名稱").Trim())
                .Where(g => !string.IsNullOrEmpty(g.Key))
                .Select(g => new
                {
                    法令名稱 = g.Key,
                    公告日 = g.OrderByDescending(r => r.Field<string>("日期")).First().Field<string>("日期"),
                    鑑別日期 = g.OrderByDescending(r => r.Field<string>("鑑別日期")).First().Field<string>("鑑別日期")
                }).ToList();

            DataTable dtShow = new DataTable();
            dtShow.Columns.Add("流水號");
            dtShow.Columns.Add("法令名稱");
            dtShow.Columns.Add("公告日");
            dtShow.Columns.Add("鑑別日期");

            int index = 1;
            foreach (var item in groupedData)
            {
                dtShow.Rows.Add(index.ToString(), item.法令名稱, item.公告日, item.鑑別日期);
                index++;
            }

            _dgvCategoryLaws.DataSource = dtShow;
            
            // 寬度設定
            if (_dgvCategoryLaws.Columns.Count > 0)
            {
                _dgvCategoryLaws.Columns[0].Width = 80;
                _dgvCategoryLaws.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // 名稱最長
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

        // 建立符合圖片紅框配色的統計表格
        private DataGridView CreateStatsGrid()
        {
            DataGridView dgv = CreateStandardGrid();
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            // 設定配色 (標題黃綠底，內容亮黃底)
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

            // 網格線顏色
            dgv.GridColor = Color.Black;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            return dgv;
        }
    }
}
