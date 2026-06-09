/// FILE: Safety_System/settings/App_DbConfig.Formula.cs ///
using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public partial class App_DbConfig
    {
        private Label _lblFStartM, _lblFStartD, _lblFEndM, _lblFEndD;

        private void BuildFormulaTab(TabPage tabFormula)
        {
            Panel pnlFormula = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

            GroupBox boxFormula = new GroupBox { Text = "資料表欄位自訂運算 (支援數學運算與區間條件判斷)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            // 🟢 Row 1: 資料庫與資料表選擇
            FlowLayoutPanel flpRow1 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 10) };
            Label lblFDb = new Label { Text = "選擇資料庫:", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaDb = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(0, 0, 30, 0) };
            _cboFormulaDb.SelectedIndexChanged += CboFormulaDb_SelectedIndexChanged;

            Label lblFTable = new Label { Text = "選擇資料表:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaTable = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaTable.SelectedIndexChanged += CboFormulaTable_SelectedIndexChanged;

            flpRow1.Controls.AddRange(new Control[] { lblFDb, _cboFormulaDb, lblFTable, _cboFormulaTable });

            // 🟢 Row 2: 對應日期欄位 與 起訖時間 (獨立下拉選單)
            FlowLayoutPanel flpRow2 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 10) };
            Label lblFMatch = new Label { Text = "對應日期欄位：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaMatchCol = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblFStart = new Label { Text = "起：", AutoSize = true, Margin = new Padding(20, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblFStartY = new Label { Text = "年", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFStartMonth = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFStartM = new Label { Text = "月", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFStartDay = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFStartD = new Label { Text = "日", AutoSize = true, Margin = new Padding(2, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };

            Label lblFEnd = new Label { Text = "迄：", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblFEndY = new Label { Text = "年", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFEndMonth = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFEndM = new Label { Text = "月", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFEndDay = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFEndD = new Label { Text = "日", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };

            InitFormulaDateComboBoxes();

            _cboFormulaMatchCol.SelectedIndexChanged += (s, e) => {
                if (_cboFormulaMatchCol.SelectedItem != null) {
                    string sel = _cboFormulaMatchCol.SelectedItem.ToString();
                    
                    bool showMonth = true;
                    bool showDay = true;

                    if (sel == "年度" || sel.EndsWith("年")) {
                        showMonth = false; showDay = false;
                    } 
                    else if (sel == "年月" || sel == "月份" || sel.EndsWith("月")) {
                        showMonth = true; showDay = false;
                    }

                    _cboFStartMonth.Visible = showMonth; _lblFStartM.Visible = showMonth;
                    _cboFEndMonth.Visible = showMonth; _lblFEndM.Visible = showMonth;
                    _cboFStartDay.Visible = showDay; _lblFStartD.Visible = showDay;
                    _cboFEndDay.Visible = showDay; _lblFEndD.Visible = showDay;
                }
            };

            Button btnClearTime = new Button { Text = "♾️ 無限區間", Width = 110, Height = 32, Margin = new Padding(15, 0, 0, 0), BackColor = Color.LightGray, Cursor = Cursors.Hand };
            btnClearTime.Click += (s, e) => {
                _cboFStartYear.SelectedIndex = 0; // 1900
                _cboFStartMonth.SelectedIndex = 0;
                // UpdateFormulaDaysCombo 會自動觸發
                _cboFStartDay.SelectedIndex = 0;
                _cboFEndYear.SelectedIndex = _cboFEndYear.Items.Count - 1; // 2099
                _cboFEndMonth.SelectedIndex = 11;
                _cboFEndDay.SelectedIndex = _cboFEndDay.Items.Count - 1;
            };

            flpRow2.Controls.AddRange(new Control[] { 
                lblFMatch, _cboFormulaMatchCol, 
                lblFStart, _cboFStartYear, lblFStartY, _cboFStartMonth, _lblFStartM, _cboFStartDay, _lblFStartD,
                lblTilde,
                lblFEnd, _cboFEndYear, lblFEndY, _cboFEndMonth, _lblFEndM, _cboFEndDay, _lblFEndD,
                btnClearTime
            });

            // 🟢 Row 3: 目標欄位 (獨立一行)
            FlowLayoutPanel flpRow3 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 10) };
            Label lblFTarget = new Label { Text = "公式結果寫入至此欄：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaTargetCol = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            flpRow3.Controls.AddRange(new Control[] { lblFTarget, _cboFormulaTargetCol });

            // 🟢 Row 4: 公式編輯區
            FlowLayoutPanel flpFormulaBlock = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(15, 10, 10, 15) };

            Label lblFormula = new Label { 
                Text = "計算公式 (如：[數量]*[單價])：\n(如需抓取浮動單價，請寫 PRICE(類別名稱)，系統將會以對應日期的最新單價進行運算)", 
                AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 10) 
            };

            FlowLayoutPanel pnlOps = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 0, 10) };
            string[] ops = { "+", "-", "*", "/", "(", ")" };
            foreach (string op in ops) {
                Button b = new Button { Text = op, Width = 45, Height = 35, Font = new Font("Consolas", 14F, FontStyle.Bold), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke };
                b.Click += (s, e) => { _rtbFormulaEditor.Focus(); _rtbFormulaEditor.SelectedText = $" {op} "; };
                pnlOps.Controls.Add(b);
            }

            Button btnInsertVar = new Button { Text = "插入欄位變數", Size = new Size(170, 35), BackColor = Color.LightSlateGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnInsertVar.FlatAppearance.BorderSize = 0;
            btnInsertVar.Click += (s, e) => {
                using(Form fSel = new Form { Text = "選擇欄位", Size = new Size(300, 400), StartPosition = FormStartPosition.CenterParent }) {
                    ListBox lb = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                    if (_cboFormulaDb.SelectedItem != null && _cboFormulaTable.SelectedItem != null) {
                        string dName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
                        string tName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;
                        var cols = DataManager.GetColumnNames(dName, tName);
                        foreach(var c in cols) lb.Items.Add(c);
                    }
                    lb.DoubleClick += (s2, e2) => {
                        if (lb.SelectedItem != null) { _rtbFormulaEditor.Focus(); _rtbFormulaEditor.SelectedText = $"[{lb.SelectedItem}]"; fSel.Close(); }
                    };
                    fSel.Controls.Add(lb); fSel.ShowDialog();
                }
            };
            pnlOps.Controls.Add(btnInsertVar);

            Button btnClearForm = new Button { Text = "✨ 清空編輯區重新設定", Size = new Size(200, 35), BackColor = Color.Gainsboro, ForeColor = Color.Black, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnClearForm.Click += (s, e) => {
                _currentFormulaEditId = 0; _rtbFormulaEditor.Clear(); _cboFormulaTargetCol.SelectedIndex = -1; _cboFormulaMatchCol.SelectedIndex = -1;
            };
            pnlOps.Controls.Add(btnClearForm);

            _rtbFormulaEditor = new RichTextBox { Width = 700, Height = 120, Font = new Font("Consolas", 14F), BackColor = Color.AliceBlue, Margin = new Padding(0, 0, 0, 15) };

            Button btnSaveFormula = new Button { Text = "💾 儲存此運算公式", Size = new Size(200, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnSaveFormula.FlatAppearance.BorderSize = 0;
            btnSaveFormula.Click += BtnSaveFormula_Click;

            flpFormulaBlock.Controls.AddRange(new Control[] { lblFormula, pnlOps, _rtbFormulaEditor, btnSaveFormula });

            boxFormula.Controls.Add(flpFormulaBlock);
            boxFormula.Controls.Add(flpRow3);
            boxFormula.Controls.Add(flpRow2);
            boxFormula.Controls.Add(flpRow1);

            TableLayoutPanel tlpListArea = new TableLayoutPanel { Dock = DockStyle.Top, Height = 400, ColumnCount = 1, RowCount = 2, Margin = new Padding(0, 20, 0, 0) };
            tlpListArea.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlpListArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            FlowLayoutPanel pnlFormulaAction = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(0, 0, 0, 5), FlowDirection = FlowDirection.LeftToRight };

            Button btnExportFormula = new Button { Text = "📤 匯出所有公式", Width = 180, Height = 40, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 15, 0) };
            btnExportFormula.FlatAppearance.BorderSize = 0;
            btnExportFormula.Click += BtnExportFormula_Click;

            Button btnImportFormula = new Button { Text = "📥 匯入公式設定", Width = 180, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0) };
            btnImportFormula.FlatAppearance.BorderSize = 0;
            btnImportFormula.Click += BtnImportFormula_Click;

            pnlFormulaAction.Controls.Add(btnExportFormula);
            pnlFormulaAction.Controls.Add(btnImportFormula);

            GroupBox boxFormulasList = new GroupBox { Text = "已設定的公式清單 (全系統) - 點擊可編輯", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            _flpFormulasList = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxFormulasList.Controls.Add(_flpFormulasList);

            tlpListArea.Controls.Add(pnlFormulaAction, 0, 0);
            tlpListArea.Controls.Add(boxFormulasList, 0, 1);

            pnlFormula.Controls.Add(tlpListArea);
            pnlFormula.Controls.Add(boxFormula);
            tabFormula.Controls.Add(pnlFormula);

            // 當 Tab 切換到公式分頁時，強制重新抓取資料繪製
            tabFormula.Enter += (s, e) => {
                RefreshAllFormulasList();
            };
        }

        // ================= 日期下拉選單管理 =================
        private void InitFormulaDateComboBoxes()
        {
            _cboFStartYear.Items.Add("1900");
            _cboFEndYear.Items.Add("1900");
            
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) {
                _cboFStartYear.Items.Add(i.ToString()); _cboFEndYear.Items.Add(i.ToString());
            }
            
            _cboFStartYear.Items.Add("2099");
            _cboFEndYear.Items.Add("2099");

            for (int i = 1; i <= 12; i++) {
                string m = i.ToString("D2");
                _cboFStartMonth.Items.Add(m); _cboFEndMonth.Items.Add(m);
            }

            _cboFStartYear.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFStartYear, _cboFStartMonth, _cboFStartDay);
            _cboFStartMonth.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFStartYear, _cboFStartMonth, _cboFStartDay);
            _cboFEndYear.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFEndYear, _cboFEndMonth, _cboFEndDay);
            _cboFEndMonth.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFEndYear, _cboFEndMonth, _cboFEndDay);

            SetFormulaDateStr(DateTime.Today.ToString("yyyy-01-01"), _cboFStartYear, _cboFStartMonth, _cboFStartDay);
            SetFormulaDateStr(DateTime.Today.ToString("yyyy-12-31"), _cboFEndYear, _cboFEndMonth, _cboFEndDay);
        }

        private void UpdateFormulaDaysCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null) return;
            if (!int.TryParse(y.SelectedItem.ToString(), out int year) || !int.TryParse(m.SelectedItem.ToString(), out int month)) return;
            
            int days = DateTime.DaysInMonth(year, month);
            string currentDay = d.SelectedItem?.ToString();
            d.Items.Clear();
            for (int i = 1; i <= days; i++) d.Items.Add(i.ToString("D2"));
            
            if (currentDay != null && d.Items.Contains(currentDay)) d.SelectedItem = currentDay;
            else d.SelectedIndex = d.Items.Count - 1;
        }

        private string GetFormulaDateStr(ComboBox y, ComboBox m, ComboBox d, string matchCol)
        {
            string year = y.SelectedItem?.ToString() ?? "1900";
            string month = m.SelectedItem?.ToString() ?? "01";
            string day = d.SelectedItem?.ToString() ?? "01";
            
            if (string.IsNullOrEmpty(matchCol)) return $"{year}-{month}-{day}";
            
            if (matchCol == "年度" || matchCol.EndsWith("年")) return year;
            if (matchCol == "年月" || matchCol == "月份" || matchCol.EndsWith("月")) return $"{year}-{month}";
            return $"{year}-{month}-{day}";
        }

        private void SetFormulaDateStr(string dateStr, ComboBox y, ComboBox m, ComboBox d)
        {
            if (string.IsNullOrEmpty(dateStr)) return;
            var parts = dateStr.Split('-');
            if (parts.Length >= 1 && y.Items.Contains(parts[0])) y.SelectedItem = parts[0];
            if (parts.Length >= 2 && m.Items.Contains(parts[1])) m.SelectedItem = parts[1];
            if (parts.Length >= 3) {
                UpdateFormulaDaysCombo(y, m, d);
                if (d.Items.Contains(parts[2])) d.SelectedItem = parts[2];
            }
        }

        // ================= 事件邏輯 =================
        private void CboFormulaDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            _cboFormulaTable.Items.Clear();
            _cboFormulaTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
            
            if (_cboFormulaDb.SelectedItem == null) return;
            var selectedDb = (ItemMap)_cboFormulaDb.SelectedItem;
            
            if (!string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                var tbItems = _dbMap[selectedDb.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                _cboFormulaTable.Items.AddRange(tbItems);
            }
        }

        private void CboFormulaTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            _cboFormulaTargetCol.Items.Clear();
            _cboFormulaMatchCol.Items.Clear();
            _cboFormulaMatchCol.Items.Add(""); 
            
            if (_cboFormulaDb.SelectedItem == null || _cboFormulaTable.SelectedItem == null) return;

            string dbName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;

            if (!string.IsNullOrEmpty(dbName) && !string.IsNullOrEmpty(tableName)) {
                var cols = DataManager.GetColumnNames(dbName, tableName).Where(c => c != "Id").ToArray();
                _cboFormulaTargetCol.Items.AddRange(cols);
                _cboFormulaMatchCol.Items.AddRange(cols);
            }
        }

        private async void BtnSaveFormula_Click(object sender, EventArgs e)
        {
            string authPrompt = "設定自動運算公式需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return; 

            if (_cboFormulaDb.SelectedItem == null || _cboFormulaTable.SelectedItem == null || _cboFormulaTargetCol.SelectedItem == null) {
                MessageBox.Show("請確認資料庫、資料表與目標欄位皆已選擇！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string dbName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;
            string targetCol = _cboFormulaTargetCol.SelectedItem.ToString();
            string matchCol = _cboFormulaMatchCol.SelectedItem?.ToString() ?? "";
            
            string sDate = GetFormulaDateStr(_cboFStartYear, _cboFStartMonth, _cboFStartDay, matchCol);
            string eDate = GetFormulaDateStr(_cboFEndYear, _cboFEndMonth, _cboFEndDay, matchCol);
            string formula = _rtbFormulaEditor.Text.Trim();

            if (string.IsNullOrEmpty(formula)) {
                MessageBox.Show("公式不可為空！若要取消該欄位的公式，請在下方清單點擊刪除。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.Compare(sDate, eDate) > 0) {
                MessageBox.Show("【起日】不能大於【迄日】！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string checkSql = "SELECT StartDate, EndDate FROM ColumnFormulas WHERE DbName=@DB AND TableName=@TB AND TargetCol=@TC AND Id != @ExId";
                    using (var cmd = new SQLiteCommand(checkSql, conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                        cmd.Parameters.AddWithValue("@TC", targetCol); cmd.Parameters.AddWithValue("@ExId", _currentFormulaEditId); 
                        using (var reader = cmd.ExecuteReader()) {
                            while (reader.Read()) {
                                string os = reader["StartDate"].ToString();
                                string oe = reader["EndDate"].ToString();
                                if (string.Compare(sDate, oe) <= 0 && string.Compare(eDate, os) >= 0) {
                                    MessageBox.Show($"此目標欄位在該時間區間內已有設定其他公式！\n重疊的區間：{os} ~ {oe}\n\n為避免資料計算衝突，請強制調整您的時間區間。", "重疊防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }

                    if (_currentFormulaEditId > 0) {
                        string updateSql = "UPDATE ColumnFormulas SET MatchCol=@MC, StartDate=@SD, EndDate=@ED, Formula=@F WHERE Id=@Id";
                        using(var cmd = new SQLiteCommand(updateSql, conn)) {
                            cmd.Parameters.AddWithValue("@MC", matchCol); cmd.Parameters.AddWithValue("@SD", sDate);
                            cmd.Parameters.AddWithValue("@ED", eDate); cmd.Parameters.AddWithValue("@F", formula);
                            cmd.Parameters.AddWithValue("@Id", _currentFormulaEditId); cmd.ExecuteNonQuery();
                        }
                    } else {
                        string insertSql = "INSERT INTO ColumnFormulas (DbName, TableName, TargetCol, MatchCol, StartDate, EndDate, Formula) VALUES (@DB, @TB, @TC, @MC, @SD, @ED, @F)";
                        using(var cmd = new SQLiteCommand(insertSql, conn)) {
                            cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                            cmd.Parameters.AddWithValue("@TC", targetCol); cmd.Parameters.AddWithValue("@MC", matchCol);
                            cmd.Parameters.AddWithValue("@SD", sDate); cmd.Parameters.AddWithValue("@ED", eDate);
                            cmd.Parameters.AddWithValue("@F", formula); cmd.ExecuteNonQuery();
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show("儲存公式時發生錯誤：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show($"【{targetCol}】 運算公式已成功儲存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            _currentFormulaEditId = 0;
            _rtbFormulaEditor.Clear();
            RefreshAllFormulasList();

            if (MessageBox.Show($"公式已儲存。\n\n是否要立即在背景重新計算【{tableName}】的所有歷史資料？\n\n(系統將以低記憶體消耗方式逐筆刷新，並自動觸發相關的資料同步)", "背景重算確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                await RunBackgroundRecalculation(dbName, tableName);
            }
        }

        private async Task RunBackgroundRecalculation(string dbName, string tableName)
        {
            bool hasChanges = false;
            
            using (ProgressForm progForm = new ProgressForm("重新計算歷史資料中..."))
            {
                await progForm.ExecuteAsync(async delegate(IProgress<int> progInt, IProgress<string> progStr)
                {
                    await Task.Run(() =>
                    {
                        progStr.Report($"正在讀取資料表：{tableName} ...");
                        progInt.Report(5);
                        
                        DataTable dt = DataManager.GetTableData(dbName, tableName, "", "", "");
                        if (dt == null || dt.Rows.Count == 0) return;

                        DataTable dtFormulas = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT * FROM ColumnFormulas WHERE DbName=@DB AND TableName=@TB", conn)) {
                                cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                                using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtFormulas);
                            }
                        }

                        if (dtFormulas.Rows.Count == 0) return;

                        System.Text.RegularExpressions.Regex priceRegex = new System.Text.RegularExpressions.Regex(@"PRICE\((?<cat>[^\)]+)\)");
                        System.Text.RegularExpressions.Regex fieldRegex = new System.Text.RegularExpressions.Regex(@"\[(.*?)\]");

                        int totalRows = dt.Rows.Count;

                        using (DataTable dtMath = new DataTable())
                        {
                            for (int i = 0; i < totalRows; i++)
                            {
                                if (i % 50 == 0 || i == totalRows - 1) {
                                    progInt.Report(5 + (int)((double)(i + 1) / totalRows * 80)); 
                                    progStr.Report($"正在計算與更新資料： 第 {i + 1} 筆 / 共 {totalRows} 筆");
                                }

                                DataRow row = dt.Rows[i];
                                bool rowChanged = false;

                                foreach (DataRow fRow in dtFormulas.Rows)
                                {
                                    string tCol = fRow["TargetCol"].ToString();
                                    string mCol = fRow["MatchCol"].ToString();
                                    string sDate = fRow["StartDate"].ToString();
                                    string eDate = fRow["EndDate"].ToString();
                                    string rawFormula = fRow["Formula"].ToString();

                                    if (!dt.Columns.Contains(tCol)) continue;

                                    if (!string.IsNullOrEmpty(mCol) && dt.Columns.Contains(mCol)) {
                                        string rDate = row[mCol]?.ToString().Trim() ?? "";
                                        if (string.IsNullOrEmpty(rDate)) continue;
                                        if (string.Compare(rDate, sDate) < 0 || string.Compare(rDate, eDate) > 0) continue;
                                    }

                                    string evalFormula = rawFormula;
                                    bool canCompute = true;

                                    if (evalFormula.Contains("PRICE(")) {
                                        DateTime targetDate = DateTime.Today;
                                        if (!string.IsNullOrEmpty(mCol) && dt.Columns.Contains(mCol)) {
                                            string dateStr = row[mCol]?.ToString().Trim() ?? "";
                                            if (dateStr.Length == 7 && dateStr.Contains("-")) DateTime.TryParse(dateStr + "-01", out targetDate);
                                            else DateTime.TryParse(dateStr, out targetDate);
                                        }

                                        var priceMatches = priceRegex.Matches(evalFormula);
                                        foreach (System.Text.RegularExpressions.Match m in priceMatches) {
                                            string category = m.Groups["cat"].Value.Trim();
                                            double unitPrice = DataManager.GetUnitPrice(category, targetDate);
                                            evalFormula = evalFormula.Replace(m.Value, unitPrice.ToString());
                                        }
                                    }

                                    var fieldMatches = fieldRegex.Matches(evalFormula);
                                    foreach (System.Text.RegularExpressions.Match m in fieldMatches) {
                                        string colName = m.Groups[1].Value;
                                        if (dt.Columns.Contains(colName)) {
                                            string val = row[colName]?.ToString().Replace(",", "").Trim();
                                            if (string.IsNullOrEmpty(val) || !double.TryParse(val, out _)) val = "0";
                                            evalFormula = evalFormula.Replace($"[{colName}]", val);
                                        } else { canCompute = false; break; }
                                    }

                                    if (canCompute) {
                                        try {
                                            object result = dtMath.Compute(evalFormula, null);
                                            if (result != DBNull.Value) {
                                                double dRes = Convert.ToDouble(result);
                                                string strRes = Math.Round(dRes, 4).ToString("0.####");
                                                if (row[tCol]?.ToString() != strRes) {
                                                    row[tCol] = strRes; rowChanged = true;
                                                }
                                            }
                                        } catch {
                                            if (row[tCol]?.ToString() != "") { row[tCol] = ""; rowChanged = true; }
                                        }
                                    }
                                }
                                if (!rowChanged) row.AcceptChanges();
                            }
                        }

                        progStr.Report("正在將計算結果寫入資料庫並觸發跨表同步...");
                        progInt.Report(90);

                        DataTable dtChanges = dt.GetChanges();
                        if (dtChanges != null && dtChanges.Rows.Count > 0) {
                            DataManager.BulkSaveTable(dbName, tableName, dtChanges, progInt, progStr);
                            hasChanges = true;
                        }
                    });
                });
            }

            if (hasChanges) MessageBox.Show("歷史資料重新計算與同步已完成！", "背景運算完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else MessageBox.Show("運算完成，所有歷史資料均為最新，無須更新。", "背景運算完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RefreshAllFormulasList()
        {
            if (_flpFormulasList == null) return;
            
            _flpFormulasList.Controls.Clear();
            DataTable dt = new DataTable();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM ColumnFormulas", conn))
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            } catch { }

            if (dt.Rows.Count == 0) {
                _flpFormulasList.Controls.Add(new Label { Text = "系統目前沒有設定任何自動運算公式。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                return;
            }

            foreach (DataRow row in dt.Rows) {
                int id = Convert.ToInt32(row["Id"]);
                string db = row["DbName"].ToString(); string tb = row["TableName"].ToString();
                string targetCol = row["TargetCol"].ToString(); string matchCol = row["MatchCol"].ToString();
                string sDate = row["StartDate"].ToString(); string eDate = row["EndDate"].ToString();
                string formula = row["Formula"].ToString();

                string dateInfo = string.IsNullOrEmpty(matchCol) ? "" : $" (當 [{matchCol}] 介於 {sDate} ~ {eDate} 時)";
                string text = $"庫:[{db}] 表:[{tb}]  ➡️  目標:[{targetCol}] = {formula}{dateInfo}";
                
                Label lTxt = new Label { Text = text, AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Cursor = Cursors.Hand };
                int reqW = TextRenderer.MeasureText(text, lTxt.Font).Width + 100;
                int panelW = Math.Max(_flpFormulasList.ClientSize.Width - 25, reqW);

                Panel p = new Panel { Width = panelW, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5), Cursor = Cursors.Hand };
                Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(panelW - 60, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                btnDel.FlatAppearance.BorderSize = 0;
                
                btnDel.Click += (s, ev) => {
                    if (MessageBox.Show($"確定刪除此自動運算公式？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        string authPrompt = "刪除運算公式需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                        if (AuthManager.VerifyAdmin(authPrompt)) {
                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var cmd = new SQLiteCommand("DELETE FROM ColumnFormulas WHERE Id=@Id", conn)) {
                                    cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery();
                                }
                            }
                            RefreshAllFormulasList();
                            _currentFormulaEditId = 0;
                        }
                    }
                };

                Action loadToEdit = () => {
                    _currentFormulaEditId = id;
                    foreach(ItemMap im in _cboFormulaDb.Items) { if (im.EnName == db) { _cboFormulaDb.SelectedItem = im; break; } }
                    foreach(ItemMap im in _cboFormulaTable.Items) { if (im.EnName == tb) { _cboFormulaTable.SelectedItem = im; break; } }
                    if (_cboFormulaTargetCol.Items.Contains(targetCol)) _cboFormulaTargetCol.SelectedItem = targetCol;
                    if (_cboFormulaMatchCol.Items.Contains(matchCol)) _cboFormulaMatchCol.SelectedItem = matchCol;
                    SetFormulaDateStr(sDate, _cboFStartYear, _cboFStartMonth, _cboFStartDay);
                    SetFormulaDateStr(eDate, _cboFEndYear, _cboFEndMonth, _cboFEndDay);
                    _rtbFormulaEditor.Text = formula;
                };

                p.Click += (s, ev) => loadToEdit();
                lTxt.Click += (s, ev) => loadToEdit();

                p.Controls.Add(lTxt);
                p.Controls.Add(btnDel);
                _flpFormulasList.Controls.Add(p);
            }
        }

        private void BtnExportFormula_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統自動運算公式設定_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT DbName AS [資料庫名], TableName AS [資料表名], TargetCol AS [目標欄位], MatchCol AS [對應日期欄位], StartDate AS [區間起日], EndDate AS [區間迄日], Formula AS [運算公式] FROM ColumnFormulas", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("自動運算公式設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("自動運算公式設定匯出成功！您可以以此檔案作為備份。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnImportFormula_Click(object sender, EventArgs e)
        {
            string authPrompt = "匯入公式設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的公式設定檔" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    new SQLiteCommand("DELETE FROM ColumnFormulas", conn, trans).ExecuteNonQuery();

                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string db = ws.Cells[r, 1].Text.Trim(); string tb = ws.Cells[r, 2].Text.Trim();
                                        string targetCol = ws.Cells[r, 3].Text.Trim(); string matchCol = ws.Cells[r, 4].Text.Trim();
                                        string sDate = ws.Cells[r, 5].Text.Trim(); string eDate = ws.Cells[r, 6].Text.Trim();
                                        string formula = ws.Cells[r, 7].Text.Trim();

                                        if (string.IsNullOrEmpty(db) || string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(targetCol) || string.IsNullOrEmpty(formula)) continue;

                                        string sql = @"INSERT INTO ColumnFormulas (DbName, TableName, TargetCol, MatchCol, StartDate, EndDate, Formula) VALUES (@DB, @TB, @TC, @MC, @SD, @ED, @F)";
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@DB", db); cmd.Parameters.AddWithValue("@TB", tb);
                                            cmd.Parameters.AddWithValue("@TC", targetCol); cmd.Parameters.AddWithValue("@MC", matchCol);
                                            cmd.Parameters.AddWithValue("@SD", sDate); cmd.Parameters.AddWithValue("@ED", eDate);
                                            cmd.Parameters.AddWithValue("@F", formula); cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("自動運算公式設定已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        RefreshAllFormulasList();
                    } catch (Exception ex) { MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }
    }
}
