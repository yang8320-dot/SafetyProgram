/// FILE: Safety_System/settings/App_ReminderManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ReminderManager : Form
    {
        private ListBox _lbRules;
        private TextBox _txtRuleName, _txtTargetUsers;
        private ComboBox _cboDb, _cboTable, _cboDateCol;
        private NumericUpDown _numAdvanceDays;
        private RichTextBox _rtbTemplate;
        private CheckBox _chkIsActive;
        
        private int _currentEditId = 0;
        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        private class ItemMap {
            public string EnName; public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        public App_ReminderManager()
        {
            ReminderEngine.InitDatabase();
            _dbMap = App_DbConfig.GetDbMapCache();
            InitializeComponent();
            LoadRulesList();
        }

        private void InitializeComponent()
        {
            this.Text = "⏰ 系統智能提醒設定";
            this.Size = new Size(1100, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // 左側清單
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            Label l1 = new Label { Text = "已建立的提醒規則", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Dock = DockStyle.Top, Height = 30 };
            _lbRules = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lbRules.SelectedIndexChanged += LbRules_SelectedIndexChanged;
            
            Button btnAdd = new Button { Text = "➕ 新增空白規則", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 10, 0, 0) };
            btnAdd.Click += (s, e) => ClearEditor();

            pnlLeft.Controls.Add(_lbRules);
            pnlLeft.Controls.Add(l1);
            pnlLeft.Controls.Add(btnAdd);

            // 右側編輯區
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
            Label l2 = new Label { Text = "編輯提醒規則條件", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkOrange, Dock = DockStyle.Top, Height = 40 };

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };

            _chkIsActive = new CheckBox { Text = "啟用此提醒規則", Checked = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.ForestGreen, Margin = new Padding(0, 0, 0, 15) };
            
            _txtRuleName = new TextBox { Width = 300 };
            _txtTargetUsers = new TextBox { Width = 300, Text = "ALL" };
            
            _cboDb = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboTable = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboDateCol = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
            
            _numAdvanceDays = new NumericUpDown { Width = 80, Minimum = 0, Maximum = 365, Value = 30 };
            _rtbTemplate = new RichTextBox { Width = 650, Height = 100, Font = new Font("Consolas", 12F), BackColor = Color.AliceBlue };

            // 資料庫連動邏輯
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
            flp.Controls.Add(BuildRow("指定接收對象：", _txtTargetUsers, "輸入登入帳號，多人用半形逗號分隔。全廠通知請填 ALL"));
            flp.Controls.Add(BuildRow("來源庫與表：", _cboDb, _cboTable));
            flp.Controls.Add(BuildRow("到期判斷日期欄位：", _cboDateCol, ""));
            flp.Controls.Add(BuildRow("提前觸發天數：", _numAdvanceDays, "天 (當 日期-今天 ≦ 此天數時觸發)"));
            
            Label lblTmpl = new Label { Text = "訊息顯示樣板 (可使用 [欄位名] 提取該筆資料)：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) };
            flp.Controls.Add(lblTmpl);
            flp.Controls.Add(_rtbTemplate);

            Button btnSave = new Button { Text = "💾 儲存規則", Width = 250, Height = 45, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 20, 0, 0) };
            btnSave.Click += BtnSave_Click;

            Button btnDel = new Button { Text = "🗑️ 刪除", Width = 120, Height = 45, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(20, 20, 0, 0) };
            btnDel.Click += BtnDel_Click;

            FlowLayoutPanel flpBtns = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            flpBtns.Controls.Add(btnSave);
            flpBtns.Controls.Add(btnDel);
            flp.Controls.Add(flpBtns);

            pnlRight.Controls.Add(flp);
            pnlRight.Controls.Add(l2);

            tlp.Controls.Add(pnlLeft, 0, 0);
            tlp.Controls.Add(pnlRight, 1, 0);
            this.Controls.Add(tlp);
        }

        // 🟢 修正 1 & 2：將參數改為 object hintOrCtrl2，解決 CS1503 與 CS8121 型別判斷錯誤
        private Panel BuildRow(string labelText, Control ctrl1, object hintOrCtrl2)
        {
            Panel p = new Panel { Width = 750, Height = 40, Margin = new Padding(0, 5, 0, 5) };
            Label l = new Label { Text = labelText, Location = new Point(0, 5), AutoSize = true };
            ctrl1.Location = new Point(170, 2);
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
            foreach (DataRow row in dt.Rows) {
                string status = Convert.ToInt32(row["IsActive"]) == 1 ? "🟢" : "⚫";
                _lbRules.Items.Add(new RuleItem { Id = Convert.ToInt32(row["Id"]), DisplayText = $"{status} {row["RuleName"]}" });
            }
        }

        private class RuleItem {
            public int Id; public string DisplayText;
            public override string ToString() => DisplayText;
        }

        private void ClearEditor()
        {
            _currentEditId = 0;
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
                
                // 🟢 修正 3：移除 AsEnumerable() 依賴，改用 foreach 防止 CS0411 編譯報錯
                foreach (DataRow r in dt.Rows) {
                    if (Convert.ToInt32(r["Id"]) == item.Id) {
                        row = r;
                        break;
                    }
                }

                if (row != null) {
                    _currentEditId = item.Id;
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

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtRuleName.Text) || _cboDb.SelectedItem == null || _cboTable.SelectedItem == null || _cboDateCol.SelectedItem == null) {
                MessageBox.Show("請確認規則名稱、資料庫、資料表與日期欄位皆已選擇！"); return;
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql;
                    if (_currentEditId == 0) {
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
                        if (_currentEditId != 0) cmd.Parameters.AddWithValue("@Id", _currentEditId);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("規則儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadRulesList();
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message); }
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (_currentEditId > 0 && MessageBox.Show("確定要刪除此規則嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM ReminderRules WHERE Id=@Id", conn)) {
                            cmd.Parameters.AddWithValue("@Id", _currentEditId); cmd.ExecuteNonQuery();
                        }
                    }
                    ClearEditor(); LoadRulesList();
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message); }
            }
        }
    }
}
