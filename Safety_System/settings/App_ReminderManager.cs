/// FILE: Safety_System/settings/App_ReminderManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public class App_ReminderManager : Form
    {
        // ================= Tab 1: 資料表條件觸發 =================
        private ComboBox _cboRuleFilter; 
        private ListBox _lbRules;
        private TextBox _txtRuleName, _txtTargetUsers;
        private ComboBox _cboDb, _cboTable, _cboDateCol;
        private NumericUpDown _numAdvanceDays;
        private RichTextBox _rtbTemplate;
        private CheckBox _chkIsActive;
        private int _currentEditRuleId = 0;

        // ================= Tab 2: 自訂待辦清單 =================
        private ComboBox _cboToDoFilter; 
        private ListBox _lbToDos;
        private TextBox _txtToDoName, _txtToDoUsers;
        private DateTimePicker _dtpToDoDate;
        private NumericUpDown _numToDoAdvance;
        private RichTextBox _rtbToDoMessage;
        private CheckBox _chkToDoIsActive;
        private CheckBox _chkIsRecurring; 
        private ComboBox _cboRecurType;   
        private int _currentEditToDoId = 0;

        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;
        private List<string> _allUsersList = new List<string>(); 

        private class ItemMap {
            public string EnName; public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        private class RuleItem {
            public int Id; public string DisplayText; public string Targets;
            public override string ToString() => DisplayText;
        }

        public App_ReminderManager()
        {
            ReminderEngine.InitDatabase();
            _dbMap = App_DbConfig.GetDbMapCache();
            LoadUsersFromDatabase();
            InitializeComponent();
            
            LoadRulesList();
            LoadToDosList();
        }

        private void LoadUsersFromDatabase()
        {
            _allUsersList.Clear();
            _allUsersList.Add("ALL");

            foreach(var u in LicenseManager.DefaultUsers) {
                if(!_allUsersList.Contains(u)) _allUsersList.Add(u);
            }

            try {
                DataTable dt = DataManager.GetTableData("SystemConfig", "AllowedUsers", "", "", "");
                if (dt != null) {
                    foreach (DataRow row in dt.Rows) {
                        string user = row["使用者帳號"].ToString().Trim();
                        if (!string.IsNullOrEmpty(user) && !_allUsersList.Contains(user)) {
                            _allUsersList.Add(user);
                        }
                    }
                }
            } catch { }
        }

        private void InitializeComponent()
        {
            this.Text = "⏰ 系統提醒設定"; 
            // 🟢 介面高度 +50
            this.Size = new Size(1100, 730);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.WhiteSmoke;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            TabControl tabMain = new TabControl { Dock = DockStyle.Fill, Padding = new Point(15, 10) };

            TabPage tabRules = new TabPage("📊 資料庫條件觸發規則") { BackColor = Color.White };
            BuildRulesTab(tabRules);

            TabPage tabToDos = new TabPage("📝 自訂待辦事項 (手動建立任務)") { BackColor = Color.White };
            BuildToDosTab(tabToDos);

            tabMain.TabPages.Add(tabRules);
            tabMain.TabPages.Add(tabToDos);

            this.Controls.Add(tabMain);
        }

        private string ShowUserSelectionDialog(string currentSelected)
        {
            string result = currentSelected;
            using (Form f = new Form { Text = "選擇接收對象", Size = new Size(350, 450), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選要收到提醒的帳號：", Dock = DockStyle.Top, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                
                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.None };
                
                var currentList = currentSelected.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();

                foreach (var u in _allUsersList) {
                    clb.Items.Add(u, currentList.Contains(u));
                }

                clb.ItemCheck += (s, e) => {
                    if (clb.Items[e.Index].ToString() == "ALL" && e.NewValue == CheckState.Checked) {
                        for (int i = 0; i < clb.Items.Count; i++) {
                            if (i != e.Index) clb.SetItemChecked(i, false);
                        }
                    } else if (clb.Items[e.Index].ToString() != "ALL" && e.NewValue == CheckState.Checked) {
                        int allIndex = clb.Items.IndexOf("ALL");
                        if (allIndex >= 0) clb.SetItemChecked(allIndex, false);
                    }
                };

                Button btnOk = new Button { Text = "確定", Dock = DockStyle.Bottom, Height = 45, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnOk.Click += (s, e) => {
                    List<string> selected = new List<string>();
                    foreach (var item in clb.CheckedItems) selected.Add(item.ToString());
                    
                    if (selected.Count == 0) result = "";
                    else result = string.Join(",", selected);
                    
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(clb);
                f.Controls.Add(btnOk);
                f.Controls.Add(lbl);

                if (f.ShowDialog() == DialogResult.OK) {
                    return result;
                }
            }
            return currentSelected;
        }

        private void BuildRulesTab(TabPage page)
        {
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 330F)); // 稍微加寬左側清單區
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // 左側清單與 IO 按鈕
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            
            Panel pnlFilter = new Panel { Dock = DockStyle.Top, Height = 60 };
            Label lFilter = new Label { Text = "篩選接收對象:", AutoSize = true, Location = new Point(0, 5), Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) };
            
            // 🟢 下拉選單 X 座標往右推，避免重疊
            _cboRuleFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200, Location = new Point(110, 3) };
            _cboRuleFilter.Items.Add("全部顯示");
            foreach (var u in _allUsersList) _cboRuleFilter.Items.Add(u);
            _cboRuleFilter.SelectedIndex = 0;
            _cboRuleFilter.SelectedIndexChanged += (s, e) => LoadRulesList();

            Label l1 = new Label { Text = "已建立的資料表提醒清單", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Location = new Point(0, 35), AutoSize = true };
            pnlFilter.Controls.Add(lFilter);
            pnlFilter.Controls.Add(_cboRuleFilter);
            pnlFilter.Controls.Add(l1);

            _lbRules = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lbRules.SelectedIndexChanged += LbRules_SelectedIndexChanged;
            
            Button btnAdd = new Button { Text = "➕ 新增空白規則", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 10, 0, 0) };
            btnAdd.Click += (s, e) => ClearRuleEditor();

            Panel pnlIo = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(0, 15, 0, 5) };
            Button btnExpRule = new Button { Text = "📤 匯出", Width = 145, Dock = DockStyle.Left, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            Button btnImpRule = new Button { Text = "📥 匯入", Width = 145, Dock = DockStyle.Right, BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            btnExpRule.Click += BtnExportRule_Click;
            btnImpRule.Click += BtnImportRule_Click;
            pnlIo.Controls.Add(btnExpRule);
            pnlIo.Controls.Add(btnImpRule);

            pnlLeft.Controls.Add(_lbRules);
            pnlLeft.Controls.Add(pnlFilter);
            pnlLeft.Controls.Add(pnlIo); 
            pnlLeft.Controls.Add(btnAdd);

            // 🟢 右側編輯區加入 AutoScroll，防止高度被裁切
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20), AutoScroll = true };
            Label l2 = new Label { Text = "編輯觸發條件與樣板", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkOrange, Dock = DockStyle.Top, Height = 40 };

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };

            // 🟢 加入 AutoSize = true
            _chkIsActive = new CheckBox { Text = "啟用此提醒規則", AutoSize = true, Checked = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.ForestGreen, Margin = new Padding(0, 0, 0, 15) };
            
            _txtRuleName = new TextBox { Width = 300 };
            
            _txtTargetUsers = new TextBox { Width = 230, Text = "ALL", ReadOnly = true, BackColor = Color.WhiteSmoke };
            Button btnPickUser = new Button { Text = "選擇人員", Width = 80, Height = 30, BackColor = Color.LightGray, Cursor = Cursors.Hand };
            btnPickUser.Click += (s, e) => {
                _txtTargetUsers.Text = ShowUserSelectionDialog(_txtTargetUsers.Text);
            };

            Panel pnlUsers = new Panel { Width = 550, Height = 35 };
            _txtTargetUsers.Location = new Point(0, 2);
            btnPickUser.Location = new Point(240, 0);
            Label lblUserHint = new Label { Text = "點擊按鈕選擇", AutoSize = true, Location = new Point(330, 5), ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 10F) };
            pnlUsers.Controls.Add(_txtTargetUsers); pnlUsers.Controls.Add(btnPickUser); pnlUsers.Controls.Add(lblUserHint);

            _cboDb = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboTable = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboDateCol = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            
            _numAdvanceDays = new NumericUpDown { Width = 80, Minimum = 0, Maximum = 365, Value = 30 };
            _rtbTemplate = new RichTextBox { Width = 650, Height = 100, Font = new Font("Consolas", 12F), BackColor = Color.AliceBlue };

            _cboDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            foreach (var kvp in _dbMap) _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
            
            _cboDb.SelectedIndexChanged += (s, e) => {
                _cboTable.Items.Clear(); _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
                var db = _cboDb.SelectedItem as ItemMap;
                if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                    foreach (var tbl in _dbMap[db.EnName].Tables) _cboTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            };

            _cboTable.SelectedIndexChanged += (s, e) => {
                _cboDateCol.Items.Clear();
                var db = _cboDb.SelectedItem as ItemMap; var tb = _cboTable.SelectedItem as ItemMap;
                if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                    var cols = DataManager.GetColumnNames(db.EnName, tb.EnName).Where(c => c != "Id" && c != "附件檔案" && c != "備註");
                    foreach (var c in cols) _cboDateCol.Items.Add(c);
                }
            };

            flp.Controls.Add(_chkIsActive);
            flp.Controls.Add(BuildRow("規則名稱：", _txtRuleName, "如：許可證屆期提醒"));
            flp.Controls.Add(BuildRow("指定接收對象：", pnlUsers, ""));
            flp.Controls.Add(BuildRow("來源庫與表：", _cboDb, _cboTable));
            flp.Controls.Add(BuildRow("到期判斷日期：", _cboDateCol, "")); 
            flp.Controls.Add(BuildRow("提前觸發天數：", _numAdvanceDays, "天 (當 日期-今天 ≦ 此天數時觸發)"));
            
            Label lblTmpl = new Label { Text = "訊息顯示樣板 (可使用 [欄位名] 提取該筆資料)：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) };
            flp.Controls.Add(lblTmpl);
            flp.Controls.Add(_rtbTemplate);

            Button btnSave = new Button { Text = "💾 儲存規則", Width = 220, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 20, 0, 0) };
            btnSave.Click += BtnSaveRule_Click;

            Button btnDel = new Button { Text = "🗑️ 刪除", Width = 150, Height = 55, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 20, 0, 0) };
            btnDel.Click += BtnDelRule_Click;

            FlowLayoutPanel flpBtns = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            flpBtns.Controls.Add(btnSave);
            flpBtns.Controls.Add(btnDel);
            flp.Controls.Add(flpBtns);

            pnlRight.Controls.Add(flp);
            pnlRight.Controls.Add(l2);

            tlp.Controls.Add(pnlLeft, 0, 0);
            tlp.Controls.Add(pnlRight, 1, 0);
            page.Controls.Add(tlp);
        }

        private void BuildToDosTab(TabPage page)
        {
            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 330F));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // 左側清單
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            
            Panel pnlFilter = new Panel { Dock = DockStyle.Top, Height = 60 };
            Label lFilter = new Label { Text = "篩選接收對象:", AutoSize = true, Location = new Point(0, 5), Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold) };
            
            // 🟢 下拉選單 X 座標往右推，避免重疊
            _cboToDoFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200, Location = new Point(110, 3) };
            _cboToDoFilter.Items.Add("全部顯示");
            foreach (var u in _allUsersList) _cboToDoFilter.Items.Add(u);
            _cboToDoFilter.SelectedIndex = 0;
            _cboToDoFilter.SelectedIndexChanged += (s, e) => LoadToDosList();

            Label l1 = new Label { Text = "已建立的待辦事項清單", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Location = new Point(0, 35), AutoSize = true };
            pnlFilter.Controls.Add(lFilter);
            pnlFilter.Controls.Add(_cboToDoFilter);
            pnlFilter.Controls.Add(l1);

            _lbToDos = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lbToDos.SelectedIndexChanged += LbToDos_SelectedIndexChanged;
            
            Button btnAdd = new Button { Text = "➕ 新增待辦事項", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.Teal, ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 10, 0, 0) };
            btnAdd.Click += (s, e) => ClearToDoEditor();

            Panel pnlIo = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(0, 15, 0, 5) };
            Button btnExpToDo = new Button { Text = "📤 匯出", Width = 135, Dock = DockStyle.Left, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            Button btnImpToDo = new Button { Text = "📥 匯入", Width = 135, Dock = DockStyle.Right, BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
            btnExpToDo.Click += BtnExportToDo_Click;
            btnImpToDo.Click += BtnImportToDo_Click;
            pnlIo.Controls.Add(btnExpToDo);
            pnlIo.Controls.Add(btnImpToDo);

            pnlLeft.Controls.Add(_lbToDos);
            pnlLeft.Controls.Add(pnlFilter);
            pnlLeft.Controls.Add(pnlIo);
            pnlLeft.Controls.Add(btnAdd);

            // 🟢 右側編輯區加入 AutoScroll
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20), AutoScroll = true };
            Label l2 = new Label { Text = "編輯待辦事項內容", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.Teal, Dock = DockStyle.Top, Height = 40 };

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };

            // 🟢 加入 AutoSize = true
            _chkToDoIsActive = new CheckBox { Text = "啟用此待辦事項", AutoSize = true, Checked = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.ForestGreen, Margin = new Padding(0, 0, 0, 15) };
            
            _txtToDoName = new TextBox { Width = 300 };
            
            _txtToDoUsers = new TextBox { Width = 230, Text = "ALL", ReadOnly = true, BackColor = Color.WhiteSmoke };
            Button btnPickUser = new Button { Text = "選擇人員", Width = 80, Height = 30, BackColor = Color.LightGray, Cursor = Cursors.Hand };
            btnPickUser.Click += (s, e) => {
                _txtToDoUsers.Text = ShowUserSelectionDialog(_txtToDoUsers.Text);
            };

            Panel pnlUsers = new Panel { Width = 550, Height = 35 };
            _txtToDoUsers.Location = new Point(0, 2);
            btnPickUser.Location = new Point(240, 0);
            Label lblUserHint = new Label { Text = "點擊按鈕選擇", AutoSize = true, Location = new Point(330, 5), ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 10F) };
            pnlUsers.Controls.Add(_txtToDoUsers); pnlUsers.Controls.Add(btnPickUser); pnlUsers.Controls.Add(lblUserHint);

            _dtpToDoDate = new DateTimePicker { Width = 150, Format = DateTimePickerFormat.Short, Font = new Font("Consolas", 13F) };
            _numToDoAdvance = new NumericUpDown { Width = 80, Minimum = 0, Maximum = 365, Value = 7 };
            _rtbToDoMessage = new RichTextBox { Width = 650, Height = 120, Font = new Font("Microsoft JhengHei UI", 12F), BackColor = Color.LightYellow };

            _chkIsRecurring = new CheckBox { Text = "🔄 這是一個循環任務", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };
            _cboRecurType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, Visible = false };
            _cboRecurType.Items.AddRange(new string[] { "每天", "每週", "每月", "每年" });
            _cboRecurType.SelectedIndex = 2; 

            _chkIsRecurring.CheckedChanged += (s, e) => {
                _cboRecurType.Visible = _chkIsRecurring.Checked;
            };

            Panel pnlRecur = new Panel { Width = 550, Height = 35 };
            _chkIsRecurring.Location = new Point(0, 5);
            _cboRecurType.Location = new Point(200, 2);
            pnlRecur.Controls.Add(_chkIsRecurring); pnlRecur.Controls.Add(_cboRecurType);

            flp.Controls.Add(_chkToDoIsActive);
            flp.Controls.Add(BuildRow("待辦標題：", _txtToDoName, "如：提交第三季環保報告"));
            flp.Controls.Add(BuildRow("指定接收對象：", pnlUsers, ""));
            flp.Controls.Add(BuildRow("首次到期基準日：", _dtpToDoDate, "")); 
            flp.Controls.Add(BuildRow("任務屬性：", pnlRecur, "設定到期後是否自動推展至下一期")); 
            flp.Controls.Add(BuildRow("提前觸發天數：", _numToDoAdvance, "天 (提早幾天開始在畫面上跳出警告)"));
            
            Label lblTmpl = new Label { Text = "詳細提醒內容 (此段文字將直接顯示給使用者看)：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) };
            flp.Controls.Add(lblTmpl);
            flp.Controls.Add(_rtbToDoMessage);

            Button btnSave = new Button { Text = "💾 儲存待辦", Width = 220, Height = 55, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 20, 0, 0) };
            btnSave.Click += BtnSaveToDo_Click;

            Button btnDel = new Button { Text = "🗑️ 刪除", Width = 150, Height = 55, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 20, 0, 0) };
            btnDel.Click += BtnDelToDo_Click;

            FlowLayoutPanel flpBtns = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            flpBtns.Controls.Add(btnSave);
            flpBtns.Controls.Add(btnDel);
            flp.Controls.Add(flpBtns);

            pnlRight.Controls.Add(flp);
            pnlRight.Controls.Add(l2);

            tlp.Controls.Add(pnlLeft, 0, 0);
            tlp.Controls.Add(pnlRight, 1, 0);
            page.Controls.Add(tlp);
        }

        // =========================================================================
        // 共用工具與邏輯
        // =========================================================================
        private Panel BuildRow(string labelText, Control ctrl1, object hintOrCtrl2)
        {
            Panel p = new Panel { Width = 750, Height = 40, Margin = new Padding(0, 5, 0, 5) };
            Label l = new Label { Text = labelText, Location = new Point(0, 5), AutoSize = true };
            
            // 🟢 控制項 X 座標大幅往右推 35px，解決嚴重重疊問題 (140 -> 175)
            ctrl1.Location = new Point(175, 2); 
            p.Controls.Add(l); p.Controls.Add(ctrl1);

            if (hintOrCtrl2 is string hint && !string.IsNullOrEmpty(hint)) {
                Label lh = new Label { Text = hint, Location = new Point(ctrl1.Right + 10, 5), AutoSize = true, ForeColor = Color.DimGray };
                p.Controls.Add(lh);
            } else if (hintOrCtrl2 is Control ctrl2) {
                ctrl2.Location = new Point(ctrl1.Right + 10, 2);
                p.Controls.Add(ctrl2);
            }
            return p;
        }

        private void LoadRulesList()
        {
            _lbRules.Items.Clear();
            DataTable dt = ReminderEngine.GetAllRules();
            
            string filterUser = _cboRuleFilter.SelectedItem?.ToString() ?? "全部顯示";

            foreach (DataRow row in dt.Rows) {
                string targets = row["TargetUsers"].ToString().Trim();
                
                if (filterUser != "全部顯示" && !targets.Equals("ALL", StringComparison.OrdinalIgnoreCase)) {
                    var usersList = targets.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(u => u.Trim()).ToList();
                    if (!usersList.Contains(filterUser, StringComparer.OrdinalIgnoreCase)) continue;
                }

                string status = Convert.ToInt32(row["IsActive"]) == 1 ? "🟢" : "⚫";
                _lbRules.Items.Add(new RuleItem { 
                    Id = Convert.ToInt32(row["Id"]), 
                    DisplayText = $"{status} {row["RuleName"]}",
                    Targets = targets
                });
            }
        }

        private void ClearRuleEditor()
        {
            _currentEditRuleId = 0;
            _txtRuleName.Clear(); _txtTargetUsers.Text = "ALL";
            _cboDb.SelectedIndex = -1; _cboTable.Items.Clear(); _cboDateCol.Items.Clear();
            _numAdvanceDays.Value = 30; _rtbTemplate.Clear(); _chkIsActive.Checked = true;
            _lbRules.ClearSelected();
        }

        private void LbRules_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lbRules.SelectedItem is RuleItem item) {
                DataTable dt = ReminderEngine.GetAllRules();
                DataRow row = null;
                
                foreach (DataRow r in dt.Rows) {
                    if (Convert.ToInt32(r["Id"]) == item.Id) {
                        row = r;
                        break;
                    }
                }

                if (row != null) {
                    _currentEditRuleId = item.Id;
                    _chkIsActive.Checked = Convert.ToInt32(row["IsActive"]) == 1;
                    _txtRuleName.Text = row["RuleName"].ToString();
                    _txtTargetUsers.Text = row["TargetUsers"].ToString();
                    
                    string db = row["DbName"].ToString();
                    string tb = row["TableName"].ToString();
                    
                    foreach (ItemMap im in _cboDb.Items) if (im.EnName == db) { _cboDb.SelectedItem = im; break; }
                    foreach (ItemMap im in _cboTable.Items) if (im.EnName == tb) { _cboTable.SelectedItem = im; break; }
                    
                    if (_cboDateCol.Items.Contains(row["DateCol"].ToString())) _cboDateCol.SelectedItem = row["DateCol"].ToString();
                    _numAdvanceDays.Value = Convert.ToInt32(row["AdvanceDays"]);
                    _rtbTemplate.Text = row["MessageTemplate"].ToString();
                }
            }
        }

        private void BtnSaveRule_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtRuleName.Text) || _cboDb.SelectedItem == null || _cboTable.SelectedItem == null || _cboDateCol.SelectedItem == null) {
                MessageBox.Show("請確認規則名稱、資料庫、資料表與日期欄位皆已選擇！"); return;
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql;
                    if (_currentEditRuleId == 0) {
                        sql = "INSERT INTO ReminderRules (RuleName, TargetUsers, DbName, TableName, DateCol, AdvanceDays, MessageTemplate, IsActive) VALUES (@RN, @TU, @DB, @TB, @DC, @AD, @MT, @IA)";
                    } else {
                        sql = "UPDATE ReminderRules SET RuleName=@RN, TargetUsers=@TU, DbName=@DB, TableName=@TB, DateCol=@DC, AdvanceDays=@AD, MessageTemplate=@MT, IsActive=@IA WHERE Id=@Id";
                    }

                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@RN", _txtRuleName.Text.Trim());
                        cmd.Parameters.AddWithValue("@TU", string.IsNullOrWhiteSpace(_txtTargetUsers.Text) ? "ALL" : _txtTargetUsers.Text.Trim());
                        cmd.Parameters.AddWithValue("@DB", ((ItemMap)_cboDb.SelectedItem).EnName);
                        cmd.Parameters.AddWithValue("@TB", ((ItemMap)_cboTable.SelectedItem).EnName);
                        cmd.Parameters.AddWithValue("@DC", _cboDateCol.SelectedItem.ToString());
                        cmd.Parameters.AddWithValue("@AD", _numAdvanceDays.Value);
                        cmd.Parameters.AddWithValue("@MT", _rtbTemplate.Text.Trim());
                        cmd.Parameters.AddWithValue("@IA", _chkIsActive.Checked ? 1 : 0);
                        if (_currentEditRuleId != 0) cmd.Parameters.AddWithValue("@Id", _currentEditRuleId);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("資料表規則儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadRulesList();
                
                // 🟢 強制觸發引擎背景掃描，確保當天新建能立刻跳出
                ReminderEngine.CheckAndShowReminders();
                
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
        }

        private void BtnDelRule_Click(object sender, EventArgs e)
        {
            if (_currentEditRuleId > 0 && MessageBox.Show("確定要刪除此規則嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM ReminderRules WHERE Id=@Id", conn)) {
                            cmd.Parameters.AddWithValue("@Id", _currentEditRuleId); cmd.ExecuteNonQuery();
                        }
                    }
                    ClearRuleEditor(); LoadRulesList();
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message); }
            }
        }

        private void BtnExportRule_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統提醒規則(資料表)_" + DateTime.Now.ToString("yyyyMMdd") })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        DataTable dt = ReminderEngine.GetAllRules();
                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("ReminderRules");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnImportRule_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的資料表規則設定檔" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;
                            
                            var headers = new Dictionary<string, int>();
                            for(int c=1; c<=ws.Dimension.Columns; c++) headers[ws.Cells[1,c].Text] = c;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    new SQLiteCommand("DELETE FROM ReminderRules", conn, trans).ExecuteNonQuery();
                                    
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string ruleName = headers.ContainsKey("RuleName") ? ws.Cells[r, headers["RuleName"]].Text : "";
                                        if (string.IsNullOrWhiteSpace(ruleName)) continue;

                                        string targetUsers = headers.ContainsKey("TargetUsers") ? ws.Cells[r, headers["TargetUsers"]].Text : "ALL";
                                        string dbName = headers.ContainsKey("DbName") ? ws.Cells[r, headers["DbName"]].Text : "";
                                        string tbName = headers.ContainsKey("TableName") ? ws.Cells[r, headers["TableName"]].Text : "";
                                        string dateCol = headers.ContainsKey("DateCol") ? ws.Cells[r, headers["DateCol"]].Text : "";
                                        string advDays = headers.ContainsKey("AdvanceDays") ? ws.Cells[r, headers["AdvanceDays"]].Text : "30";
                                        string msgTemplate = headers.ContainsKey("MessageTemplate") ? ws.Cells[r, headers["MessageTemplate"]].Text : "";
                                        string isActive = headers.ContainsKey("IsActive") ? ws.Cells[r, headers["IsActive"]].Text : "1";

                                        string sql = "INSERT INTO ReminderRules (RuleName, TargetUsers, DbName, TableName, DateCol, AdvanceDays, MessageTemplate, IsActive) VALUES (@RN, @TU, @DB, @TB, @DC, @AD, @MT, @IA)";
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@RN", ruleName);
                                            cmd.Parameters.AddWithValue("@TU", targetUsers);
                                            cmd.Parameters.AddWithValue("@DB", dbName);
                                            cmd.Parameters.AddWithValue("@TB", tbName);
                                            cmd.Parameters.AddWithValue("@DC", dateCol);
                                            cmd.Parameters.AddWithValue("@AD", string.IsNullOrEmpty(advDays) ? 30 : int.Parse(advDays));
                                            cmd.Parameters.AddWithValue("@MT", msgTemplate);
                                            cmd.Parameters.AddWithValue("@IA", string.IsNullOrEmpty(isActive) ? 1 : int.Parse(isActive));
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("設定已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadRulesList();
                    } catch (Exception ex) {
                        MessageBox.Show("匯入失敗，請確認檔案格式：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // --- Tab 2 (自訂待辦清單) 邏輯 ---
        private void LoadToDosList()
        {
            _lbToDos.Items.Clear();
            DataTable dt = ReminderEngine.GetAllToDos();
            string filterUser = _cboToDoFilter.SelectedItem?.ToString() ?? "全部顯示";

            foreach (DataRow row in dt.Rows) {
                string targets = row["TargetUsers"].ToString().Trim();
                
                if (filterUser != "全部顯示" && !targets.Equals("ALL", StringComparison.OrdinalIgnoreCase)) {
                    var usersList = targets.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(u => u.Trim()).ToList();
                    if (!usersList.Contains(filterUser, StringComparer.OrdinalIgnoreCase)) continue;
                }

                string status = Convert.ToInt32(row["IsActive"]) == 1 ? "🟢" : "⚫";
                string recurMark = (row.Table.Columns.Contains("IsRecurring") && row["IsRecurring"] != DBNull.Value && Convert.ToInt32(row["IsRecurring"]) == 1) ? "🔄" : "";
                
                _lbToDos.Items.Add(new RuleItem { Id = Convert.ToInt32(row["Id"]), DisplayText = $"{status}{recurMark} {row["TaskName"]}" });
            }
        }

        private void ClearToDoEditor()
        {
            _currentEditToDoId = 0;
            _txtToDoName.Clear(); _txtToDoUsers.Text = "ALL";
            _dtpToDoDate.Value = DateTime.Today;
            _numToDoAdvance.Value = 7; _rtbToDoMessage.Clear(); _chkToDoIsActive.Checked = true;
            _chkIsRecurring.Checked = false; _cboRecurType.SelectedIndex = 2; 
            _lbToDos.ClearSelected();
        }

        private void LbToDos_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lbToDos.SelectedItem is RuleItem item) {
                DataTable dt = ReminderEngine.GetAllToDos();
                DataRow row = null;
                foreach (DataRow r in dt.Rows) {
                    if (Convert.ToInt32(r["Id"]) == item.Id) { row = r; break; }
                }

                if (row != null) {
                    _currentEditToDoId = item.Id;
                    _chkToDoIsActive.Checked = Convert.ToInt32(row["IsActive"]) == 1;
                    _txtToDoName.Text = row["TaskName"].ToString();
                    _txtToDoUsers.Text = row["TargetUsers"].ToString();
                    
                    if (DateTime.TryParse(row["DueDate"].ToString(), out DateTime d)) _dtpToDoDate.Value = d;
                    else _dtpToDoDate.Value = DateTime.Today;

                    _numToDoAdvance.Value = Convert.ToInt32(row["AdvanceDays"]);
                    _rtbToDoMessage.Text = row["Message"].ToString();

                    if (row.Table.Columns.Contains("IsRecurring") && row["IsRecurring"] != DBNull.Value) {
                        _chkIsRecurring.Checked = Convert.ToInt32(row["IsRecurring"]) == 1;
                    } else {
                        _chkIsRecurring.Checked = false;
                    }

                    if (row.Table.Columns.Contains("RecurType")) {
                        string rType = row["RecurType"].ToString();
                        if (_cboRecurType.Items.Contains(rType)) _cboRecurType.SelectedItem = rType;
                    }
                }
            }
        }

        private void BtnSaveToDo_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtToDoName.Text) || string.IsNullOrWhiteSpace(_rtbToDoMessage.Text)) {
                MessageBox.Show("請確認待辦標題與提醒內容皆已填寫！"); return;
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql;
                    if (_currentEditToDoId == 0) {
                        sql = "INSERT INTO CustomToDos (TaskName, TargetUsers, DueDate, AdvanceDays, Message, IsActive, IsRecurring, RecurType) VALUES (@TN, @TU, @DD, @AD, @M, @IA, @IR, @RT)";
                    } else {
                        sql = "UPDATE CustomToDos SET TaskName=@TN, TargetUsers=@TU, DueDate=@DD, AdvanceDays=@AD, Message=@M, IsActive=@IA, IsRecurring=@IR, RecurType=@RT WHERE Id=@Id";
                    }

                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@TN", _txtToDoName.Text.Trim());
                        cmd.Parameters.AddWithValue("@TU", string.IsNullOrWhiteSpace(_txtToDoUsers.Text) ? "ALL" : _txtToDoUsers.Text.Trim());
                        cmd.Parameters.AddWithValue("@DD", _dtpToDoDate.Value.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@AD", _numToDoAdvance.Value);
                        cmd.Parameters.AddWithValue("@M", _rtbToDoMessage.Text.Trim());
                        cmd.Parameters.AddWithValue("@IA", _chkToDoIsActive.Checked ? 1 : 0);
                        cmd.Parameters.AddWithValue("@IR", _chkIsRecurring.Checked ? 1 : 0);
                        cmd.Parameters.AddWithValue("@RT", _cboRecurType.Text);
                        
                        if (_currentEditToDoId != 0) cmd.Parameters.AddWithValue("@Id", _currentEditToDoId);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("自訂待辦事項儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadToDosList();

                // 🟢 強制觸發引擎背景掃描，確保當天新建能立刻跳出
                ReminderEngine.CheckAndShowReminders();

            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
        }

        private void BtnDelToDo_Click(object sender, EventArgs e)
        {
            if (_currentEditToDoId > 0 && MessageBox.Show("確定要刪除此待辦事項嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM CustomToDos WHERE Id=@Id", conn)) {
                            cmd.Parameters.AddWithValue("@Id", _currentEditToDoId); cmd.ExecuteNonQuery();
                        }
                    }
                    ClearToDoEditor(); LoadToDosList();
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message); }
            }
        }

        private void BtnExportToDo_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統待辦清單設定_" + DateTime.Now.ToString("yyyyMMdd") })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        DataTable dt = ReminderEngine.GetAllToDos();
                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("CustomToDos");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnImportToDo_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的待辦清單檔" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;
                            
                            var headers = new Dictionary<string, int>();
                            for(int c=1; c<=ws.Dimension.Columns; c++) headers[ws.Cells[1,c].Text] = c;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    new SQLiteCommand("DELETE FROM CustomToDos", conn, trans).ExecuteNonQuery();
                                    
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string taskName = headers.ContainsKey("TaskName") ? ws.Cells[r, headers["TaskName"]].Text : "";
                                        if (string.IsNullOrWhiteSpace(taskName)) continue;

                                        string targetUsers = headers.ContainsKey("TargetUsers") ? ws.Cells[r, headers["TargetUsers"]].Text : "ALL";
                                        string dueDate = headers.ContainsKey("DueDate") ? ws.Cells[r, headers["DueDate"]].Text : DateTime.Today.ToString("yyyy-MM-dd");
                                        string advDays = headers.ContainsKey("AdvanceDays") ? ws.Cells[r, headers["AdvanceDays"]].Text : "7";
                                        string message = headers.ContainsKey("Message") ? ws.Cells[r, headers["Message"]].Text : "";
                                        string isActive = headers.ContainsKey("IsActive") ? ws.Cells[r, headers["IsActive"]].Text : "1";
                                        string isRecur = headers.ContainsKey("IsRecurring") ? ws.Cells[r, headers["IsRecurring"]].Text : "0";
                                        string rType = headers.ContainsKey("RecurType") ? ws.Cells[r, headers["RecurType"]].Text : "";

                                        string sql = "INSERT INTO CustomToDos (TaskName, TargetUsers, DueDate, AdvanceDays, Message, IsActive, IsRecurring, RecurType) VALUES (@TN, @TU, @DD, @AD, @M, @IA, @IR, @RT)";
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@TN", taskName);
                                            cmd.Parameters.AddWithValue("@TU", targetUsers);
                                            cmd.Parameters.AddWithValue("@DD", dueDate);
                                            cmd.Parameters.AddWithValue("@AD", string.IsNullOrEmpty(advDays) ? 7 : int.Parse(advDays));
                                            cmd.Parameters.AddWithValue("@M", message);
                                            cmd.Parameters.AddWithValue("@IA", string.IsNullOrEmpty(isActive) ? 1 : int.Parse(isActive));
                                            cmd.Parameters.AddWithValue("@IR", string.IsNullOrEmpty(isRecur) ? 0 : int.Parse(isRecur));
                                            cmd.Parameters.AddWithValue("@RT", rType);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("待辦事項已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadToDosList();
                    } catch (Exception ex) {
                        MessageBox.Show("匯入失敗，請確認檔案格式：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
