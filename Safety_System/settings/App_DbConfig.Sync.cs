/// FILE: Safety_System/settings/App_DbConfig.Sync.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml; // 🟢 引入 EPPlus 匯出入功能

namespace Safety_System
{
    public partial class App_DbConfig
    {
        private void BuildSyncTab(TabPage tabSync)
        {
            Panel pnlSync = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };
            
            GroupBox boxSync = new GroupBox { Text = "資料同步設定 (來源儲存時自動聚合計算至目標表)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };

            TableLayoutPanel tlpSync = new TableLayoutPanel {
                Dock = DockStyle.Top, AutoSize = true, ColumnCount = 10, Padding = new Padding(15, 20, 15, 10)
            };
            
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40F)); 
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 12F)); 

            string[] headers = { "【來源】庫", "來源表", "比對欄(如:日期)", "要同步之欄位", "➡️", "【目標】庫", "寫入目標表", "比對欄(如:年月)", "接收寫入之欄位", "同步方向" };
            for(int i=0; i<10; i++) {
                tlpSync.Controls.Add(new Label { Text = headers[i], TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) }, i, 0);
            }

            for (int i = 0; i < 5; i++)
            {
                var rowUi = new SyncRowUI();
                rowUi.CboSrcDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSrcTable = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSrcMatchCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSrcSyncCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };

                Label lblArrow = new Label { Text = "➡️", TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };

                rowUi.CboTgtDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboTgtTable = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboTgtMatchCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboTgtSyncCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList }; 
                
                rowUi.CboSyncType = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSyncType.Items.AddRange(new string[] { "", "單向同步", "雙向同步" });
                rowUi.CboSyncType.SelectedIndex = 0;

                BindSyncRowEvents(rowUi);

                tlpSync.Controls.Add(rowUi.CboSrcDb, 0, i+1);
                tlpSync.Controls.Add(rowUi.CboSrcTable, 1, i+1);
                tlpSync.Controls.Add(rowUi.CboSrcMatchCol, 2, i+1);
                tlpSync.Controls.Add(rowUi.CboSrcSyncCol, 3, i+1);
                tlpSync.Controls.Add(lblArrow, 4, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtDb, 5, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtTable, 6, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtMatchCol, 7, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtSyncCol, 8, i+1);
                tlpSync.Controls.Add(rowUi.CboSyncType, 9, i+1);

                _syncRows.Add(rowUi);
            }

            boxSync.Controls.Add(tlpSync);

            Panel pnlSyncBottom = new Panel { Dock = DockStyle.Bottom, Height = 140, Padding = new Padding(15) };
            Button btnSaveSync = new Button { Text = "儲存新增的資料同步設定", Location = new Point(15, 10), Size = new Size(250, 45), BackColor = Color.Teal, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveSync.Click += BtnSaveSync_Click;
            
            Button btnShowSyncRulesList = new Button { Text = "同步設定清單", Location = new Point(290, 10), Size = new Size(180, 45), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnShowSyncRulesList.Click += BtnShowSyncRulesList_Click;

            Label lblSyncInfo = new Label { Text = "※ 僅需填入要啟用的列即可，空列系統自動忽略。若目標表無該接收欄位，系統會自動新增。\n※ 同步支援【雙向同步】(限1:1對應)，若比對欄位不一致(如 日期 vs 年月)，為防呆將強制設為單向聚合加總。", Location = new Point(15, 65), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };
            
            pnlSyncBottom.Controls.Add(btnSaveSync);
            pnlSyncBottom.Controls.Add(btnShowSyncRulesList);
            pnlSyncBottom.Controls.Add(lblSyncInfo);
            boxSync.Controls.Add(pnlSyncBottom);

            pnlSync.Controls.Add(boxSync);
            tabSync.Controls.Add(pnlSync);
        }

        private void BindSyncRowEvents(SyncRowUI r)
        {
            Action checkSyncType = () => {
                string srcM = r.CboSrcMatchCol.Text;
                string tgtM = r.CboTgtMatchCol.Text;
                if (!string.IsNullOrEmpty(srcM) && !string.IsNullOrEmpty(tgtM)) {
                    if (srcM != tgtM || (srcM.Contains("日") && tgtM.Contains("月"))) {
                        r.CboSyncType.SelectedIndex = 1; 
                        r.CboSyncType.Enabled = false;
                    } else {
                        r.CboSyncType.Enabled = true;
                    }
                }
            };

            r.CboSrcDb.SelectedIndexChanged += (s, e) => {
                r.CboSrcTable.Items.Clear(); r.CboSrcMatchCol.Items.Clear(); r.CboSrcSyncCol.Items.Clear();
                r.CboSrcTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
                if (r.CboSrcDb.SelectedItem != null) {
                    var map = (ItemMap)r.CboSrcDb.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && _dbMap.ContainsKey(map.EnName)) {
                        var tbItems = _dbMap[map.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                        r.CboSrcTable.Items.AddRange(tbItems);
                    }
                }
            };

            r.CboSrcTable.SelectedIndexChanged += (s, e) => {
                r.CboSrcMatchCol.Items.Clear(); r.CboSrcSyncCol.Items.Clear();
                r.CboSrcMatchCol.Items.Add(""); r.CboSrcSyncCol.Items.Add("");
                if (r.CboSrcTable.SelectedItem != null) {
                    var map = (ItemMap)r.CboSrcDb.SelectedItem;
                    var tmap = (ItemMap)r.CboSrcTable.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && !string.IsNullOrEmpty(tmap.EnName)) {
                        var cols = DataManager.GetColumnNames(map.EnName, tmap.EnName).Where(c => c != "Id").ToArray();
                        r.CboSrcMatchCol.Items.AddRange(cols); r.CboSrcSyncCol.Items.AddRange(cols);
                    }
                }
            };

            r.CboTgtDb.SelectedIndexChanged += (s, e) => {
                r.CboTgtTable.Items.Clear(); r.CboTgtMatchCol.Items.Clear(); r.CboTgtSyncCol.Items.Clear();
                r.CboTgtTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
                if (r.CboTgtDb.SelectedItem != null) {
                    var map = (ItemMap)r.CboTgtDb.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && _dbMap.ContainsKey(map.EnName)) {
                        var tbItems = _dbMap[map.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                        r.CboTgtTable.Items.AddRange(tbItems);
                    }
                }
            };

            r.CboTgtTable.SelectedIndexChanged += (s, e) => {
                r.CboTgtMatchCol.Items.Clear(); r.CboTgtSyncCol.Items.Clear();
                r.CboTgtSyncCol.DropDownStyle = ComboBoxStyle.DropDown;
                r.CboTgtMatchCol.Items.Add(""); r.CboTgtSyncCol.Items.Add("");
                if (r.CboTgtTable.SelectedItem != null) {
                    var map = (ItemMap)r.CboTgtDb.SelectedItem;
                    var tmap = (ItemMap)r.CboTgtTable.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && !string.IsNullOrEmpty(tmap.EnName)) {
                        var cols = DataManager.GetColumnNames(map.EnName, tmap.EnName).Where(c => c != "Id").ToArray();
                        r.CboTgtMatchCol.Items.AddRange(cols); r.CboTgtSyncCol.Items.AddRange(cols);
                    }
                }
            };

            r.CboSrcMatchCol.TextChanged += (s, e) => checkSyncType();
            r.CboTgtMatchCol.TextChanged += (s, e) => checkSyncType();
        }

        private void BtnSaveSync_Click(object sender, EventArgs e)
        {
            HashSet<string> requiredMenusToUnlock = new HashSet<string>();
            foreach (var r in _syncRows) 
            {
                if (r.CboSrcDb.SelectedItem == null || r.CboTgtDb.SelectedItem == null) continue;
                string srcDb = ((ItemMap)r.CboSrcDb.SelectedItem).EnName;
                string tgtDb = ((ItemMap)r.CboTgtDb.SelectedItem).EnName;
                if (string.IsNullOrEmpty(srcDb) || string.IsNullOrEmpty(tgtDb)) continue;
                if (srcDb == "Menu1DB" || tgtDb == "Menu1DB") requiredMenusToUnlock.Add("選單1");
                if (srcDb == "Menu2DB" || tgtDb == "Menu2DB") requiredMenusToUnlock.Add("選單2");
                if (srcDb == "Menu3DB" || tgtDb == "Menu3DB") requiredMenusToUnlock.Add("選單3");
                if (srcDb == "Menu4DB" || tgtDb == "Menu4DB") requiredMenusToUnlock.Add("選單4");
            }

            foreach (string menu in requiredMenusToUnlock) {
                if (!AuthManager.VerifyHiddenMenu(menu)) {
                    MessageBox.Show($"已取消儲存！因為您未通過【{menu}】的密碼驗證。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; 
                }
            }

            try {
                string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                HashSet<string> existingRules = new HashSet<string>();
                using (var connCheck = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                    connCheck.Open();
                    using (var cmdCheck = new SQLiteCommand("SELECT * FROM SyncRules", connCheck))
                    using (var reader = cmdCheck.ExecuteReader()) {
                        while (reader.Read()) {
                            string key = $"{reader["SrcDb"]}|{reader["SrcTable"]}|{reader["SrcMatchCol"]}|{reader["SrcSyncCol"]}|{reader["TgtDb"]}|{reader["TgtTable"]}|{reader["TgtMatchCol"]}|{reader["TgtSyncCol"]}";
                            existingRules.Add(key);
                        }
                    }
                }

                using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        foreach (var r in _syncRows) {
                            if (r.CboSrcDb.SelectedItem == null || r.CboSrcTable.SelectedItem == null || string.IsNullOrWhiteSpace(r.CboSrcMatchCol.Text) || string.IsNullOrWhiteSpace(r.CboSrcSyncCol.Text) ||
                                r.CboTgtDb.SelectedItem == null || r.CboTgtTable.SelectedItem == null || string.IsNullOrWhiteSpace(r.CboTgtMatchCol.Text) || string.IsNullOrWhiteSpace(r.CboTgtSyncCol.Text))
                                continue;

                            string srcDb = ((ItemMap)r.CboSrcDb.SelectedItem).EnName;
                            string srcTbl = ((ItemMap)r.CboSrcTable.SelectedItem).EnName;
                            string tgtDb = ((ItemMap)r.CboTgtDb.SelectedItem).EnName;
                            string tgtTbl = ((ItemMap)r.CboTgtTable.SelectedItem).EnName;

                            if (string.IsNullOrEmpty(srcDb) || string.IsNullOrEmpty(tgtDb)) continue;

                            string currentKey = $"{srcDb}|{srcTbl}|{r.CboSrcMatchCol.Text}|{r.CboSrcSyncCol.Text}|{tgtDb}|{tgtTbl}|{r.CboTgtMatchCol.Text}|{r.CboTgtSyncCol.Text}";
                            
                            if (existingRules.Contains(currentKey)) {
                                MessageBox.Show($"偵測到重複的同步設定！\n來源表：【{((ItemMap)r.CboSrcTable.SelectedItem).ChName}】\n目標表：【{((ItemMap)r.CboTgtTable.SelectedItem).ChName}】\n請勿重複新增！", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                trans.Rollback(); return; 
                            }
                            existingRules.Add(currentKey);

                            string sql = "INSERT INTO SyncRules (SrcDb, SrcTable, SrcMatchCol, SrcSyncCol, TgtDb, TgtTable, TgtMatchCol, TgtSyncCol, SyncType) VALUES (@SD, @ST, @SMC, @SSC, @TD, @TT, @TMC, @TSC, @Type)";
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@SD", srcDb); cmd.Parameters.AddWithValue("@ST", srcTbl);
                                cmd.Parameters.AddWithValue("@SMC", r.CboSrcMatchCol.Text); cmd.Parameters.AddWithValue("@SSC", r.CboSrcSyncCol.Text);
                                cmd.Parameters.AddWithValue("@TD", tgtDb); cmd.Parameters.AddWithValue("@TT", tgtTbl);
                                cmd.Parameters.AddWithValue("@TMC", r.CboTgtMatchCol.Text); cmd.Parameters.AddWithValue("@TSC", r.CboTgtSyncCol.Text);
                                cmd.Parameters.AddWithValue("@Type", string.IsNullOrEmpty(r.CboSyncType.Text) ? "單向同步" : r.CboSyncType.Text);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                MessageBox.Show("新增的資料同步設定已成功儲存至清單！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                foreach (var r in _syncRows) {
                    r.CboSrcDb.SelectedIndex = 0; r.CboTgtDb.SelectedIndex = 0;
                    r.CboSrcMatchCol.Text = ""; r.CboSrcSyncCol.Text = "";
                    r.CboTgtMatchCol.Text = ""; r.CboTgtSyncCol.Text = "";
                }
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BtnShowSyncRulesList_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "同步設定清單", Size = new Size(1000, 600), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.Sizable, MaximizeBox = true, MinimizeBox = false, BackColor = Color.White })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F)); 
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                Panel pnlTop = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15, 15, 15, 5) };
                Label lbl = new Label { Text = "已啟用之跨表同步規則清單：", Dock = DockStyle.Left, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
                
                Button btnExp = new Button { Text = "📤 匯出", Dock = DockStyle.Right, Width = 100, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button btnImp = new Button { Text = "📥 匯入", Dock = DockStyle.Right, Width = 100, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Margin = new Padding(0,0,10,0) };
                
                pnlTop.Controls.Add(lbl);
                pnlTop.Controls.Add(btnImp);
                pnlTop.Controls.Add(btnExp);

                tlp.Controls.Add(pnlTop, 0, 0);

                FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10) };
                tlp.Controls.Add(flp, 0, 1);
                f.Controls.Add(tlp);

                Action loadRules = null;
                loadRules = () => {
                    flp.Controls.Clear();
                    try {
                        string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT * FROM SyncRules", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }
                        if (dt.Rows.Count == 0) {
                            flp.Controls.Add(new Label { Text = "目前沒有任何同步設定。", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) });
                            return;
                        }
                        foreach (DataRow row in dt.Rows) {
                            int id = Convert.ToInt32(row["Id"]);
                            string sDb = row["SrcDb"].ToString(); string sTb = row["SrcTable"].ToString(); string sMatch = row["SrcMatchCol"].ToString(); string sSync = row["SrcSyncCol"].ToString();
                            string tDb = row["TgtDb"].ToString(); string tTb = row["TgtTable"].ToString(); string tMatch = row["TgtMatchCol"].ToString(); string tSync = row["TgtSyncCol"].ToString();
                            string type = row.Table.Columns.Contains("SyncType") ? row["SyncType"].ToString() : "單向同步";

                            string text = $"【{type}】 {sDb}.{sTb}[{sMatch} ➔ {sSync}]  ➡️  {tDb}.{tTb}[{tMatch} ➔ {tSync}]";
                            Label lTxt = new Label { Text = text, AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F) };
                            
                            int reqW = TextRenderer.MeasureText(text, lTxt.Font).Width + 100;
                            int panelW = Math.Max(flp.ClientSize.Width - 25, reqW);

                            Panel p = new Panel { Width = panelW, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5) };
                            Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(panelW - 60, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                            btnDel.FlatAppearance.BorderSize = 0;
                            btnDel.Click += (s, ev) => {
                                if (MessageBox.Show($"確定刪除此同步規則？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                                    using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                                        conn.Open();
                                        using (var cmd = new SQLiteCommand("DELETE FROM SyncRules WHERE Id=@Id", conn)) {
                                            cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery();
                                        }
                                    }
                                    loadRules();
                                }
                            };
                            p.Controls.Add(lTxt); p.Controls.Add(btnDel); flp.Controls.Add(p);
                        }
                    } catch { }
                };

                // 🟢 修正：更改變數名稱避免與外層的 e 衝突
                btnExp.Click += (senderObj, ev) => {
                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "資料同步規則_" + DateTime.Now.ToString("yyyyMMdd") }) {
                        if (sfd.ShowDialog() == DialogResult.OK) {
                            try {
                                DataTable dt = new DataTable();
                                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                    conn.Open();
                                    using (var cmd = new SQLiteCommand("SELECT SrcDb AS [來源庫], SrcTable AS [來源表], SrcMatchCol AS [來源比對欄], SrcSyncCol AS [來源同步欄], TgtDb AS [目標庫], TgtTable AS [目標表], TgtMatchCol AS [目標比對欄], TgtSyncCol AS [目標接收欄], SyncType AS [同步方向] FROM SyncRules", conn))
                                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                                }
                                using (ExcelPackage p = new ExcelPackage()) {
                                    var ws = p.Workbook.Worksheets.Add("同步設定");
                                    ws.Cells["A1"].LoadFromDataTable(dt, true);
                                    ws.Cells.AutoFitColumns();
                                    p.SaveAs(new FileInfo(sfd.FileName));
                                }
                                MessageBox.Show("匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        }
                    }
                };

                // 🟢 修正：更改變數名稱避免與外層的 e 衝突，並將 loadKeys() 改為 loadRules()
                btnImp.Click += (senderObj, ev) => {
                    string authPrompt = "匯入同步設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                    if (!AuthManager.VerifyAdmin(authPrompt)) return;

                    using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的設定檔" }) {
                        if (ofd.ShowDialog() == DialogResult.OK) {
                            try {
                                using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                                    ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                    if (ws == null || ws.Dimension == null) return;

                                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                        conn.Open();
                                        using (var trans = conn.BeginTransaction()) {
                                            for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                                string sDb = ws.Cells[r, 1].Text.Trim();
                                                string sTb = ws.Cells[r, 2].Text.Trim();
                                                string sMatch = ws.Cells[r, 3].Text.Trim();
                                                string sSync = ws.Cells[r, 4].Text.Trim();
                                                string tDb = ws.Cells[r, 5].Text.Trim();
                                                string tTb = ws.Cells[r, 6].Text.Trim();
                                                string tMatch = ws.Cells[r, 7].Text.Trim();
                                                string tSync = ws.Cells[r, 8].Text.Trim();
                                                string type = ws.Cells[r, 9].Text.Trim();

                                                if (string.IsNullOrEmpty(sDb) || string.IsNullOrEmpty(tDb) || string.IsNullOrEmpty(sMatch) || string.IsNullOrEmpty(tMatch)) continue;

                                                string sql = "INSERT INTO SyncRules (SrcDb, SrcTable, SrcMatchCol, SrcSyncCol, TgtDb, TgtTable, TgtMatchCol, TgtSyncCol, SyncType) VALUES (@SD, @ST, @SMC, @SSC, @TD, @TT, @TMC, @TSC, @Type)";
                                                using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                                    cmd.Parameters.AddWithValue("@SD", sDb); cmd.Parameters.AddWithValue("@ST", sTb);
                                                    cmd.Parameters.AddWithValue("@SMC", sMatch); cmd.Parameters.AddWithValue("@SSC", sSync);
                                                    cmd.Parameters.AddWithValue("@TD", tDb); cmd.Parameters.AddWithValue("@TT", tTb);
                                                    cmd.Parameters.AddWithValue("@TMC", tMatch); cmd.Parameters.AddWithValue("@TSC", tSync);
                                                    cmd.Parameters.AddWithValue("@Type", string.IsNullOrEmpty(type) ? "單向同步" : type);
                                                    cmd.ExecuteNonQuery();
                                                }
                                            }
                                            trans.Commit();
                                        }
                                    }
                                }
                                MessageBox.Show("同步設定已批次匯入並寫入成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                loadRules(); // 🟢 修正為 loadRules()
                            } catch (Exception ex) { MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        }
                    }
                };

                loadRules();
                f.ShowDialog();
            }
        }
    }
}
