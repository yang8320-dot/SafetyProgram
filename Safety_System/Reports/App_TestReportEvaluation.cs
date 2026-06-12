/// FILE: Safety_System/Reports/App_TestReportEvaluation.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_TestReportEvaluation
    {
        // 框1: 操作區控制項
        private ComboBox _cboDb;
        private ComboBox _cboTable;
        private ComboBox _cboDate;
        private ComboBox _cboPoint;
        
        // 框2: 表單區控制項
        private DateTimePicker _dtpEvalDate;
        private ComboBox _cboCompliance;
        private TextBox _txtTestDate;
        private TextBox _txtTestName;
        private TextBox _txtTestPurpose;
        private RichTextBox _rtbAnalysis;
        private DataGridView _dgvItems;

        private const string EvalDbName = "TestData";
        private const string EvalTableName = "TestReportEvaluations";
        
        // 目前載入的紀錄 ID (0代表新增)
        private int _currentId = 0;

        private class ItemMap {
            public string EnName; public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        public Control GetView()
        {
            // 初始化評估紀錄表
            string schema = TableSchemaManager.SchemaMap.ContainsKey(EvalTableName) ? TableSchemaManager.SchemaMap[EvalTableName] : "[資料庫] TEXT, [資料表] TEXT, [測定日期] TEXT, [檢測名稱] TEXT, [評估日期] TEXT, [符合度] TEXT, [測定用途] TEXT, [分析與結果說明] TEXT, [最後修改人] TEXT, [修改時間] TEXT";
            DataManager.InitTable(EvalDbName, EvalTableName, $"CREATE TABLE IF NOT EXISTS [{EvalTableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});");

            Panel mainScrollPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AutoScroll = true, Padding = new Padding(20) };
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 2 };

            // ==========================================
            // 第一個框：資料選擇與操作列
            // ==========================================
            GroupBox box1 = new GroupBox { Text = "⚙️ 檢測資料載入與操作區", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15), Margin = new Padding(0,0,0,20) };
            
            FlowLayoutPanel flpRow1 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0,5,0,10) };
            _cboDb = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboDate = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboPoint = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Button btnLoadData = new Button { Text = "📥 載入數據", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnLoadData.FlatAppearance.BorderSize = 0;
            btnLoadData.Click += BtnLoadData_Click;

            flpRow1.Controls.AddRange(new Control[] {
                new Label { Text = "資料庫:", AutoSize = true, Margin = new Padding(10,5,5,0) }, _cboDb,
                new Label { Text = "資料表:", AutoSize = true, Margin = new Padding(15,5,5,0) }, _cboTable,
                new Label { Text = "測定日期:", AutoSize = true, Margin = new Padding(15,5,5,0) }, _cboDate,
                new Label { Text = "檢測名稱(點):", AutoSize = true, Margin = new Padding(15,5,5,0) }, _cboPoint,
                new Panel { Width=10, Height=1 }, btnLoadData
            });

            FlowLayoutPanel flpRow2 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0,5,0,0) };
            
            // 🟢 修正：統一按鈕尺寸為 140x40，並對齊 Margin 解決遮擋與下移問題
            Button btnHistory = new Button { Text = "📂 讀取報告", Size = new Size(140, 40), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 10, 0) };
            btnHistory.FlatAppearance.BorderSize = 0;
            btnHistory.Click += BtnHistory_Click;

            Button btnSave = new Button { Text = "💾 存檔 / 覆蓋", Size = new Size(140, 40), BackColor = Color.ForestGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 10, 0) };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;

            Button btnWord = new Button { Text = "📝 導出 Word", Size = new Size(140, 40), BackColor = Color.RoyalBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 10, 0) };
            btnWord.FlatAppearance.BorderSize = 0;
            btnWord.Click += BtnWord_Click;

            Button btnPdf = new Button { Text = "📄 導出 PDF", Size = new Size(140, 40), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 10, 0) };
            btnPdf.FlatAppearance.BorderSize = 0;
            btnPdf.Click += BtnPdf_Click;

            flpRow2.Controls.AddRange(new Control[] { btnHistory, btnSave, btnWord, btnPdf });

            box1.Controls.Add(flpRow2);
            box1.Controls.Add(flpRow1);
            layout.Controls.Add(box1, 0, 0);

            // ==========================================
            // 第二個框：報表模板
            // ==========================================
            GroupBox box2 = new GroupBox { Text = "📄 檢測報告分析評估表 (預覽與編輯區)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            Panel pnlReport = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(20) };
            
            TableLayoutPanel tlpForm = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single };
            tlpForm.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tlpForm.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlpForm.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tlpForm.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            
            // 行 1
            tlpForm.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            tlpForm.Controls.Add(CreateHeaderLabel("評估日期："), 0, 0);
            _dtpEvalDate = new DateTimePicker { Format = DateTimePickerFormat.Short, Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(5, 8, 5, 5) };
            tlpForm.Controls.Add(_dtpEvalDate, 1, 0);
            
            tlpForm.Controls.Add(CreateHeaderLabel("符合度："), 2, 0);
            _cboCompliance = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(5, 8, 5, 5) };
            _cboCompliance.Items.AddRange(new string[] { "符合", "不符合" });
            _cboCompliance.SelectedIndex = 0;
            tlpForm.Controls.Add(_cboCompliance, 3, 0);

            // 行 2
            tlpForm.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            tlpForm.Controls.Add(CreateHeaderLabel("測定日期："), 0, 1);
            _txtTestDate = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White, Margin = new Padding(5, 10, 5, 5) };
            tlpForm.Controls.Add(_txtTestDate, 1, 1);
            tlpForm.SetColumnSpan(_txtTestDate, 3);

            // 行 3
            tlpForm.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            tlpForm.Controls.Add(CreateHeaderLabel("檢測名稱："), 0, 2);
            _txtTestName = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White, Margin = new Padding(5, 10, 5, 5) };
            tlpForm.Controls.Add(_txtTestName, 1, 2);
            tlpForm.SetColumnSpan(_txtTestName, 3);

            // 行 4
            tlpForm.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            tlpForm.Controls.Add(CreateHeaderLabel("測定用途："), 0, 3);
            _txtTestPurpose = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(5, 8, 5, 5) };
            tlpForm.Controls.Add(_txtTestPurpose, 1, 3);
            tlpForm.SetColumnSpan(_txtTestPurpose, 3);

            // 列表 (🟢 高度修改為 350)
            _dgvItems = new DataGridView {
                Dock = DockStyle.Top, 
                Height = 350, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                ReadOnly = false, // 🟢 允許編輯 (為了 CheckBox)
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                RowHeadersVisible = false, 
                Font = new Font("Microsoft JhengHei UI", 11F),
                Margin = new Padding(0, 0, 0, 0), 
                CellBorderStyle = DataGridViewCellBorderStyle.Single, 
                GridColor = Color.Black
            };
            _dgvItems.EnableHeadersVisualStyles = false;
            _dgvItems.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            _dgvItems.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            _dgvItems.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dgvItems.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            _dgvItems.ColumnHeadersHeight = 40;
            _dgvItems.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            // 🟢 新增 Checkbox 欄位提供勾選匯出功能
            DataGridViewCheckBoxColumn chkCol = new DataGridViewCheckBoxColumn { Name = "匯出", HeaderText = "匯出", Width = 60, AutoSizeMode = DataGridViewAutoSizeColumnMode.None };
            _dgvItems.Columns.Add(chkCol);

            _dgvItems.Columns.Add("項目", "項目");
            _dgvItems.Columns.Add("管制值", "管制值");
            _dgvItems.Columns.Add("測定方法", "測定方法");
            _dgvItems.Columns.Add("前次測值", "前次測值");
            _dgvItems.Columns.Add("本次測值", "本次測值");
            _dgvItems.Columns.Add("備註", "備註");

            // 設定除 Checkbox 外皆為唯讀
            foreach (DataGridViewColumn col in _dgvItems.Columns) {
                if (col.Name != "匯出") col.ReadOnly = true;
            }

            // 🟢 修改排版，確保 RichTextBox 邊框明顯，且不會覆蓋標籤
            Panel pnlAnalysis = new Panel { Dock = DockStyle.Top, Height = 220, Padding = new Padding(0, 10, 0, 0) };
            Label lblAnalysis = new Label { Text = "分析與結果說明：", Dock = DockStyle.Top, Height = 30, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            _rtbAnalysis = new RichTextBox { 
                Dock = DockStyle.Fill, 
                BorderStyle = BorderStyle.FixedSingle, 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                BackColor = Color.LightYellow // 給一個顯眼的背景色，讓使用者清楚這是一個輸入框
            };
            
            pnlAnalysis.Controls.Add(_rtbAnalysis);
            pnlAnalysis.Controls.Add(lblAnalysis);

            Label lblFooter = new Label { Text = "8-ES-B11-01", Dock = DockStyle.Bottom, Height = 30, TextAlign = ContentAlignment.BottomLeft, Font = new Font("Microsoft JhengHei UI", 10F) };

            pnlReport.Controls.Add(lblFooter);
            pnlReport.Controls.Add(pnlAnalysis);
            pnlReport.Controls.Add(_dgvItems);
            pnlReport.Controls.Add(tlpForm);

            box2.Controls.Add(pnlReport);
            layout.Controls.Add(box2, 0, 1);

            mainScrollPanel.Controls.Add(layout);

            InitDropdowns();

            return mainScrollPanel;
        }

        private Label CreateHeaderLabel(string text)
        {
            return new Label { Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), BackColor = Color.WhiteSmoke };
        }

        private void InitDropdowns()
        {
            var dbMap = App_DbConfig.GetDbMapCache();

            _cboDb.SelectedIndexChanged += (s, e) => {
                _cboTable.Items.Clear();
                _cboDate.Items.Clear();
                _cboPoint.Items.Clear();
                if (_cboDb.SelectedItem == null) return;
                
                string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
                if (dbMap.ContainsKey(dbName)) {
                    foreach (var tb in dbMap[dbName].Tables) {
                        if (tb.Key != EvalTableName) { 
                            _cboTable.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                        }
                    }
                }
                if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
            };

            _cboTable.SelectedIndexChanged += (s, e) => {
                _cboDate.Items.Clear();
                _cboPoint.Items.Clear();
                if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) return;
                
                string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
                string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;
                
                try {
                    DataTable dt = DataManager.GetTableData(dbName, tbName, "", "", "");
                    string dateCol = GetDateColumnName(dt);
                    string pointCol = GetPointColumnName(dt);
                    
                    if (!string.IsNullOrEmpty(dateCol) && !string.IsNullOrEmpty(pointCol)) {
                        var dates = new HashSet<string>();
                        foreach (DataRow r in dt.Rows) {
                            if (r[dateCol] != DBNull.Value && !string.IsNullOrEmpty(r[dateCol].ToString())) {
                                dates.Add(r[dateCol].ToString());
                            }
                        }
                        var sortedDates = new List<string>(dates);
                        sortedDates.Sort();
                        sortedDates.Reverse(); 
                        foreach(var d in sortedDates) _cboDate.Items.Add(d);
                    }
                } catch { }

                if (_cboDate.Items.Count > 0) _cboDate.SelectedIndex = 0;
            };

            _cboDate.SelectedIndexChanged += (s, e) => {
                _cboPoint.Items.Clear();
                if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null || _cboDate.SelectedItem == null) return;
                
                string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
                string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;
                string dateStr = _cboDate.SelectedItem.ToString();

                try {
                    DataTable dt = DataManager.GetTableData(dbName, tbName, "", "", "");
                    string dateCol = GetDateColumnName(dt);
                    string pointCol = GetPointColumnName(dt);

                    if (!string.IsNullOrEmpty(dateCol) && !string.IsNullOrEmpty(pointCol)) {
                        var points = new HashSet<string>();
                        foreach (DataRow r in dt.Rows) {
                            if (r[dateCol].ToString() == dateStr && r[pointCol] != DBNull.Value && !string.IsNullOrEmpty(r[pointCol].ToString())) {
                                points.Add(r[pointCol].ToString());
                            }
                        }
                        foreach(var p in points) _cboPoint.Items.Add(p);
                    }
                } catch { }

                if (_cboPoint.Items.Count > 0) _cboPoint.SelectedIndex = 0;
            };

            _cboDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            foreach (var kvp in dbMap) {
                _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
            }

            int idx = 0;
            for (int i = 0; i < _cboDb.Items.Count; i++) {
                if (((ItemMap)_cboDb.Items[i]).EnName == "TestData") { idx = i; break; }
            }
            if (_cboDb.Items.Count > 0) _cboDb.SelectedIndex = idx;
        }

        private string GetDateColumnName(DataTable dt)
        {
            if (dt.Columns.Contains("日期")) return "日期";
            if (dt.Columns.Contains("年月")) return "年月";
            if (dt.Columns.Contains("清運日期")) return "清運日期";
            return "";
        }

        private string GetPointColumnName(DataTable dt)
        {
            List<string> possibleNames = new List<string> { "檢測點", "檢測名稱", "設備名稱", "名稱", "點位", "SEG編號", "水錶名稱", "化學物質名稱", "項目", "單位" };
            foreach (string pc in possibleNames) {
                if (dt.Columns.Contains(pc)) return pc;
            }
            return "";
        }

        private string GetItemColumnName(DataTable dt)
        {
            List<string> possibleNames = new List<string> { "檢測項目", "項目", "檢測點", "檢測名稱", "名稱" };
            foreach (string pc in possibleNames) {
                if (dt.Columns.Contains(pc)) return pc;
            }
            return "";
        }

        private void BtnLoadData_Click(object sender, EventArgs e)
        {
            if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null || _cboDate.SelectedItem == null || _cboPoint.SelectedItem == null) {
                MessageBox.Show("請確認資料庫、資料表、日期與檢測名稱(點)皆已選擇！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            _currentId = 0; 
            _dtpEvalDate.Value = DateTime.Today;
            _cboCompliance.SelectedIndex = 0;
            _txtTestPurpose.Clear();
            _rtbAnalysis.Clear();

            string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
            string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;
            string dateStr = _cboDate.SelectedItem.ToString();
            string pointName = _cboPoint.SelectedItem.ToString();

            _txtTestDate.Text = dateStr;
            _txtTestName.Text = pointName;
            _dgvItems.Rows.Clear();

            try {
                DataTable dt = DataManager.GetTableData(dbName, tbName, "", "", "");
                if (dt == null) return;

                string dateCol = GetDateColumnName(dt);
                string pointCol = GetPointColumnName(dt);
                string itemCol = GetItemColumnName(dt);

                if (string.IsNullOrEmpty(dateCol) || string.IsNullOrEmpty(pointCol)) return;

                DataView dv = new DataView(dt);
                dv.RowFilter = $"[{dateCol}] = '{dateStr}' AND [{pointCol}] = '{pointName}'";

                foreach (DataRowView drv in dv) 
                {
                    string item = !string.IsNullOrEmpty(itemCol) ? drv[itemCol].ToString() : "";
                    string limit = drv.Row.Table.Columns.Contains("管制值") ? drv["管制值"].ToString() : "";
                    string method = drv.Row.Table.Columns.Contains("測定方法") ? drv["測定方法"].ToString() : "";
                    string currVal = drv.Row.Table.Columns.Contains("檢測數據") ? drv["檢測數據"].ToString() : "";
                    string note = drv.Row.Table.Columns.Contains("備註") ? drv["備註"].ToString() : "";

                    string prevVal = "N/A";
                    
                    if (!string.IsNullOrEmpty(item)) {
                        DataView dvPrev = new DataView(dt);
                        dvPrev.RowFilter = $"[{pointCol}] = '{pointName}' AND [{itemCol}] = '{item}' AND [{dateCol}] < '{dateStr}'";
                        dvPrev.Sort = $"[{dateCol}] DESC";
                        if (dvPrev.Count > 0 && dvPrev[0].Row.Table.Columns.Contains("檢測數據")) {
                            prevVal = dvPrev[0]["檢測數據"].ToString();
                        }
                    }

                    // 🟢 載入時預設打勾 true
                    _dgvItems.Rows.Add(true, item, limit, method, prevVal, currVal, note);
                }
            } 
            catch (Exception ex) {
                MessageBox.Show("載入明細失敗：" + ex.Message);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtTestDate.Text) || string.IsNullOrEmpty(_txtTestName.Text)) {
                MessageBox.Show("請先載入檢測數據後再進行存檔！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
            string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;

            DataTable dtEval = DataManager.GetTableData(EvalDbName, EvalTableName, "", "", "");
            DataRow targetRow = null;

            foreach (DataRow r in dtEval.Rows) {
                if (r.RowState != DataRowState.Deleted &&
                    r["資料庫"].ToString() == dbName &&
                    r["資料表"].ToString() == tbName &&
                    r["測定日期"].ToString() == _txtTestDate.Text &&
                    r["檢測名稱"].ToString() == _txtTestName.Text) 
                {
                    targetRow = r;
                    break;
                }
            }

            if (targetRow == null) {
                targetRow = dtEval.NewRow();
                dtEval.Rows.Add(targetRow);
            }

            targetRow["資料庫"] = dbName;
            targetRow["資料表"] = tbName;
            targetRow["測定日期"] = _txtTestDate.Text;
            targetRow["檢測名稱"] = _txtTestName.Text;
            targetRow["評估日期"] = _dtpEvalDate.Value.ToString("yyyy-MM-dd");
            targetRow["符合度"] = _cboCompliance.SelectedItem.ToString();
            targetRow["測定用途"] = _txtTestPurpose.Text;
            targetRow["分析與結果說明"] = _rtbAnalysis.Text;

            if (DataManager.BulkSaveTable(EvalDbName, EvalTableName, dtEval)) {
                MessageBox.Show("評估報告儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                DataTable dtCheck = DataManager.GetTableData(EvalDbName, EvalTableName, "", "", "");
                foreach (DataRow r in dtCheck.Rows) {
                    if (r["資料庫"].ToString() == dbName && r["資料表"].ToString() == tbName && r["測定日期"].ToString() == _txtTestDate.Text && r["檢測名稱"].ToString() == _txtTestName.Text) {
                        _currentId = Convert.ToInt32(r["Id"]); break;
                    }
                }
            } else {
                MessageBox.Show("儲存失敗！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnHistory_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "📂 歷史評估報告查詢", Size = new Size(1100, 600), StartPosition = FormStartPosition.CenterParent })
            {
                DataTable dt = DataManager.GetTableData(EvalDbName, EvalTableName, "", "", "");
                if (dt != null && dt.Columns.Contains("Id")) dt.Columns["Id"].ReadOnly = true;

                DataGridView dgv = new DataGridView { 
                    Dock = DockStyle.Fill, BackgroundColor = Color.WhiteSmoke, AllowUserToAddRows = false, ReadOnly = true,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect, RowHeadersVisible = false, Font = new Font("Microsoft JhengHei UI", 11F),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                };
                dgv.DataSource = dt;
                if (dgv.Columns.Contains("Id")) dgv.Columns["Id"].Visible = false;

                Button btnLoad = new Button { Text = "✔️ 載入選定報告", Dock = DockStyle.Bottom, Height = 45, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                Button btnDel = new Button { Text = "❌ 刪除選定報告", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };

                btnLoad.Click += (s, ev) => {
                    if (dgv.SelectedRows.Count > 0) {
                        var row = dgv.SelectedRows[0];
                        string dbName = row.Cells["資料庫"].Value?.ToString() ?? "TestData";
                        string tbName = row.Cells["資料表"].Value.ToString();
                        string tDate = row.Cells["測定日期"].Value.ToString();
                        string tName = row.Cells["檢測名稱"].Value.ToString();

                        foreach (ItemMap item in _cboDb.Items) {
                            if (item.EnName == dbName) { _cboDb.SelectedItem = item; break; }
                        }
                        
                        foreach (ItemMap item in _cboTable.Items) {
                            if (item.EnName == tbName) { _cboTable.SelectedItem = item; break; }
                        }

                        _cboDate.Items.Clear(); _cboDate.Items.Add(tDate); _cboDate.SelectedIndex = 0;
                        _cboPoint.Items.Clear(); _cboPoint.Items.Add(tName); _cboPoint.SelectedIndex = 0;

                        BtnLoadData_Click(null, null);

                        _currentId = Convert.ToInt32(row.Cells["Id"].Value);
                        if (DateTime.TryParse(row.Cells["評估日期"].Value.ToString(), out DateTime dtVal)) _dtpEvalDate.Value = dtVal;
                        _cboCompliance.SelectedItem = row.Cells["符合度"].Value.ToString();
                        _txtTestPurpose.Text = row.Cells["測定用途"].Value.ToString();
                        _rtbAnalysis.Text = row.Cells["分析與結果說明"].Value.ToString();

                        f.DialogResult = DialogResult.OK;
                    }
                };

                btnDel.Click += (s, ev) => {
                    if (dgv.SelectedRows.Count > 0 && MessageBox.Show("確定刪除此評估報告？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                        int id = Convert.ToInt32(dgv.SelectedRows[0].Cells["Id"].Value);
                        DataManager.DeleteRecord(EvalDbName, EvalTableName, id);
                        dgv.DataSource = DataManager.GetTableData(EvalDbName, EvalTableName, "", "", "");
                    }
                };

                dgv.CellDoubleClick += (s, ev) => btnLoad.PerformClick();

                f.Controls.Add(dgv);
                f.Controls.Add(btnLoad);
                f.Controls.Add(btnDel);
                f.ShowDialog();
            }
        }

        // ====================================================================
        // 🟢 PDF 導出系統 (完美套用台玻標準模版)
        // ====================================================================
        // 用來追蹤分頁列印的變數
        private int _pdfPrintRowIndex = 0;
        private int _pdfPrintPageNumber = 1;
        private bool _pdfPrintedFormInfo = false;

        private void BtnPdf_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtTestDate.Text)) {
                MessageBox.Show("請先載入報告內容！"); return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = $"檢測報告分析評估表_{_txtTestName.Text}_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _pdfPrintRowIndex = 0;
                    _pdfPrintPageNumber = 1;
                    _pdfPrintedFormInfo = false;

                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = false; 
                    pd.DefaultPageSettings.Margins = new Margins(40, 40, 50, 50);

                    pd.PrintPage += DrawReportPage;

                    try {
                        pd.Print();
                        MessageBox.Show("PDF 導出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("PDF 導出失敗：" + ex.Message, "錯誤");
                    }
                }
            }
        }

        private void DrawReportPage(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            float w = e.MarginBounds.Width;

            Font fMainTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
            Font fSubTitle = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
            Font fSign = new Font("Microsoft JhengHei UI", 12F);
            Font fDate = new Font("Microsoft JhengHei UI", 11F);
            
            Font fHeader = new Font("Microsoft JhengHei UI", 12F);
            Font fGridHead = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);
            Font fGridBody = new Font("Microsoft JhengHei UI", 11F);

            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
            StringFormat sfTopLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Near };

            // ================= 1. 大標題與簽核 (每頁都有) =================
            g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fMainTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
            y += 40;
            g.DrawString("檢測報告分析評估表", fSubTitle, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
            y += 40;
            string sign = "廠主管：______________    經/副理：______________    課/股長：______________    制表：______________";
            g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
            y += 40;

            float rowH = 35;
            
            // ================= 2. 檢測基本資料 (只在第一頁畫) =================
            if (!_pdfPrintedFormInfo)
            {
                float labelW = 120;
                float valW1 = (w - labelW * 2) / 2;
                
                RectangleF r1 = new RectangleF(x, y, labelW, rowH);
                RectangleF r2 = new RectangleF(x + labelW, y, valW1, rowH);
                RectangleF r3 = new RectangleF(x + labelW + valW1, y, labelW, rowH);
                RectangleF r4 = new RectangleF(x + labelW * 2 + valW1, y, valW1, rowH);
                
                g.FillRectangle(Brushes.WhiteSmoke, r1); g.DrawRectangle(Pens.Black, Rectangle.Round(r1)); g.DrawString("評估日期：", fHeader, Brushes.Black, r1, sfCenter);
                g.DrawRectangle(Pens.Black, Rectangle.Round(r2)); g.DrawString("  " + _dtpEvalDate.Value.ToString("yyyy/MM/dd"), fHeader, Brushes.Black, r2, sfLeft);
                g.FillRectangle(Brushes.WhiteSmoke, r3); g.DrawRectangle(Pens.Black, Rectangle.Round(r3)); g.DrawString("符合度：", fHeader, Brushes.Black, r3, sfCenter);
                g.DrawRectangle(Pens.Black, Rectangle.Round(r4)); g.DrawString("  " + _cboCompliance.Text, fHeader, Brushes.Black, r4, sfLeft);
                y += rowH;

                RectangleF rDateL = new RectangleF(x, y, labelW, rowH);
                RectangleF rDateV = new RectangleF(x + labelW, y, w - labelW, rowH);
                g.FillRectangle(Brushes.WhiteSmoke, rDateL); g.DrawRectangle(Pens.Black, Rectangle.Round(rDateL)); g.DrawString("測定日期：", fHeader, Brushes.Black, rDateL, sfCenter);
                g.DrawRectangle(Pens.Black, Rectangle.Round(rDateV)); g.DrawString("  " + _txtTestDate.Text, fHeader, Brushes.Black, rDateV, sfLeft);
                y += rowH;

                RectangleF rNameL = new RectangleF(x, y, labelW, rowH);
                RectangleF rNameV = new RectangleF(x + labelW, y, w - labelW, rowH);
                g.FillRectangle(Brushes.WhiteSmoke, rNameL); g.DrawRectangle(Pens.Black, Rectangle.Round(rNameL)); g.DrawString("檢測名稱：", fHeader, Brushes.Black, rNameL, sfCenter);
                g.DrawRectangle(Pens.Black, Rectangle.Round(rNameV)); g.DrawString("  " + _txtTestName.Text, fHeader, Brushes.Black, rNameV, sfLeft);
                y += rowH;

                RectangleF rPurpL = new RectangleF(x, y, labelW, rowH);
                RectangleF rPurpV = new RectangleF(x + labelW, y, w - labelW, rowH);
                g.FillRectangle(Brushes.WhiteSmoke, rPurpL); g.DrawRectangle(Pens.Black, Rectangle.Round(rPurpL)); g.DrawString("測定用途：", fHeader, Brushes.Black, rPurpL, sfCenter);
                g.DrawRectangle(Pens.Black, Rectangle.Round(rPurpV)); g.DrawString("  " + _txtTestPurpose.Text, fHeader, Brushes.Black, rPurpV, sfLeft);
                y += rowH;

                _pdfPrintedFormInfo = true;
            }

            // ================= 3. 資料清單 =================
            float[] colWidths = { w * 0.15f, w * 0.15f, w * 0.25f, w * 0.15f, w * 0.15f, w * 0.15f };
            float currX = x;

            // 畫 Grid 標題列
            for (int i = 1; i < _dgvItems.Columns.Count; i++) { // 從 index 1 開始，跳過 CheckBox
                RectangleF rHead = new RectangleF(currX, y, colWidths[i - 1], rowH);
                g.FillRectangle(Brushes.LightGray, rHead);
                g.DrawRectangle(Pens.Black, Rectangle.Round(rHead));
                g.DrawString(_dgvItems.Columns[i].HeaderText, fGridHead, Brushes.Black, rHead, sfCenter);
                currX += colWidths[i - 1];
            }
            y += rowH;

            // 畫 Grid 內容列
            while (_pdfPrintRowIndex < _dgvItems.Rows.Count) 
            {
                DataGridViewRow dgvRow = _dgvItems.Rows[_pdfPrintRowIndex];
                
                // 🟢 檢查勾選狀態，若沒勾則跳過此列
                if (!Convert.ToBoolean(dgvRow.Cells["匯出"].Value)) {
                    _pdfPrintRowIndex++;
                    continue;
                }

                currX = x;
                float maxH = rowH;
                for (int i = 1; i < _dgvItems.Columns.Count; i++) {
                    string val = dgvRow.Cells[i].Value?.ToString() ?? "";
                    SizeF sz = g.MeasureString(val, fGridBody, (int)colWidths[i - 1], sfCenter);
                    if (sz.Height + 10 > maxH) maxH = sz.Height + 10;
                }

                // 檢查是否超出底線 (預留分析說明的空間約 100px 以及頁碼 30px)
                if (y + maxH > e.MarginBounds.Bottom - 130) {
                    // 本頁畫不下了，直接分頁
                    g.DrawString($"第 {_pdfPrintPageNumber} 頁", fDate, Brushes.Black, new RectangleF(x, e.MarginBounds.Bottom - 15, w, 20), sfCenter);
                    _pdfPrintPageNumber++;
                    e.HasMorePages = true;
                    return;
                }

                for (int i = 1; i < _dgvItems.Columns.Count; i++) {
                    RectangleF rCell = new RectangleF(currX, y, colWidths[i - 1], maxH);
                    g.DrawRectangle(Pens.Black, Rectangle.Round(rCell));
                    string val = dgvRow.Cells[i].Value?.ToString() ?? "";
                    g.DrawString(val, fGridBody, Brushes.Black, rCell, sfCenter);
                    currX += colWidths[i - 1];
                }
                y += maxH;
                _pdfPrintRowIndex++;
            }

            // ================= 4. 分析與結果說明 =================
            float remainingH = e.MarginBounds.Bottom - 30 - y;
            if (remainingH < 80) {
                // 如果底部空間太小，就把分析說明擠到下一頁
                g.DrawString($"第 {_pdfPrintPageNumber} 頁", fDate, Brushes.Black, new RectangleF(x, e.MarginBounds.Bottom - 15, w, 20), sfCenter);
                _pdfPrintPageNumber++;
                e.HasMorePages = true;
                return;
            }

            RectangleF rAnalysis = new RectangleF(x, y, w, remainingH);
            g.DrawRectangle(Pens.Black, Rectangle.Round(rAnalysis));
            g.DrawString("分析與結果說明：\n\n" + _rtbAnalysis.Text, fHeader, Brushes.Black, new RectangleF(x + 10, y + 10, w - 20, remainingH - 20), sfTopLeft);

            // ================= 5. 底部代碼與頁碼 =================
            g.DrawString("8-ES-B11-01", fGridBody, Brushes.Black, x, e.MarginBounds.Bottom - 20);
            g.DrawString($"第 {_pdfPrintPageNumber} 頁", fDate, Brushes.Black, new RectangleF(x, e.MarginBounds.Bottom - 15, w, 20), sfCenter);

            e.HasMorePages = false;
        }

        // ====================================================================
        // 🟢 Word 導出系統 (使用 HTML，精準對齊 PDF 排版，加入勾選過濾)
        // ====================================================================
        private void BtnWord_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtTestDate.Text)) {
                MessageBox.Show("請先載入報告內容！"); return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Word 檔案 (*.docx)|*.docx", FileName = $"檢測報告分析評估表_{_txtTestName.Text}_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("<html><head><meta charset='utf-8'><style>");
                        sb.AppendLine("body { font-family: '微軟正黑體', '標楷體', sans-serif; }");
                        sb.AppendLine("table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }");
                        sb.AppendLine("th, td { border: 1px solid black; padding: 8px; text-align: center; }");
                        sb.AppendLine(".info-table td { text-align: left; }");
                        sb.AppendLine("</style></head><body>");
                        
                        // 模擬 PDF 的頂部排版
                        sb.AppendLine("<h2 style='text-align:center; margin-bottom:5px;'>台灣玻璃工業股份有限公司 - 彰濱廠</h2>");
                        sb.AppendLine("<h3 style='text-align:center; margin-top:0px;'>檢測報告分析評估表</h3>");
                        sb.AppendLine("<p style='text-align:center;'>廠主管：______________&nbsp;&nbsp;&nbsp;&nbsp;經/副理：______________&nbsp;&nbsp;&nbsp;&nbsp;課/股長：______________&nbsp;&nbsp;&nbsp;&nbsp;制表：______________</p>");
                        
                        sb.AppendLine("<hr style='margin-bottom:20px;' />");

                        sb.AppendLine("<table class='info-table'>");
                        sb.AppendLine($"<tr><td width='20%' style='background-color:#F5F5F5;'><b>評估日期：</b></td><td width='30%'>{_dtpEvalDate.Value:yyyy/MM/dd}</td><td width='20%' style='background-color:#F5F5F5;'><b>符合度：</b></td><td width='30%'>{_cboCompliance.Text}</td></tr>");
                        sb.AppendLine($"<tr><td style='background-color:#F5F5F5;'><b>測定日期：</b></td><td colspan='3'>{_txtTestDate.Text}</td></tr>");
                        sb.AppendLine($"<tr><td style='background-color:#F5F5F5;'><b>檢測名稱：</b></td><td colspan='3'>{_txtTestName.Text}</td></tr>");
                        sb.AppendLine($"<tr><td style='background-color:#F5F5F5;'><b>測定用途：</b></td><td colspan='3'>{_txtTestPurpose.Text}</td></tr>");
                        sb.AppendLine("</table>");

                        sb.AppendLine("<table>");
                        sb.AppendLine("<tr style='background-color:#D3D3D3;'><th>項目</th><th>管制值</th><th>測定方法</th><th>前次測值</th><th>本次測值</th><th>備註</th></tr>");
                        
                        // 🟢 檢查 CheckBox 狀態，只匯出有勾選的列
                        foreach (DataGridViewRow row in _dgvItems.Rows) {
                            if (!Convert.ToBoolean(row.Cells["匯出"].Value)) continue;

                            sb.AppendLine("<tr>");
                            for (int i = 1; i < _dgvItems.Columns.Count; i++) { // 跳過索引 0 的 checkbox
                                sb.AppendLine($"<td>{(row.Cells[i].Value ?? "")}</td>");
                            }
                            sb.AppendLine("</tr>");
                        }
                        sb.AppendLine("</table>");

                        sb.AppendLine("<div style='border: 1px solid black; padding: 10px; min-height: 200px;'>");
                        sb.AppendLine("<b>分析與結果說明：</b><br/>");
                        string analysis = _rtbAnalysis.Text.Replace("\n", "<br/>");
                        sb.AppendLine(analysis);
                        sb.AppendLine("</div>");

                        sb.AppendLine("<p style='text-align:left; margin-top:20px;'>8-ES-B11-01</p>");
                        
                        sb.AppendLine("</body></html>");

                        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        MessageBox.Show("Word 導出成功！\n(備註：此檔案為 HTML 封裝，若開啟時 Word 提示檔案格式問題，請點擊「是」即可正常檢視與編輯。)", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("Word 導出失敗：" + ex.Message, "錯誤");
                    }
                }
            }
        }
    }
}
