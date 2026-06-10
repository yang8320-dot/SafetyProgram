/// FILE: Safety_System/ReminderEngine.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public static class ReminderEngine
    {
        private const string DbName = "SystemConfig";
        private const string RulesTable = "ReminderRules";
        private const string LogsTable = "UserReminderLogs";
        private const string ToDosTable = "CustomToDos"; 

        public class TriggeredReminder
        {
            public int RuleId { get; set; }
            public int RecordId { get; set; }
            public string RuleName { get; set; }
            public string Message { get; set; }
            public int DaysLeft { get; set; }
        }

        public static void InitDatabase()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql1 = $@"CREATE TABLE IF NOT EXISTS [{RulesTable}] (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, RuleName TEXT, TargetUsers TEXT, DbName TEXT, 
                        TableName TEXT, DateCol TEXT, AdvanceDays INTEGER, MessageTemplate TEXT, IsActive INTEGER DEFAULT 1);";
                    string sql2 = $@"CREATE TABLE IF NOT EXISTS [{LogsTable}] (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, UserName TEXT, RuleId INTEGER, RecordId INTEGER, NextRemindDate TEXT);";
                    
                    string sql3 = $@"CREATE TABLE IF NOT EXISTS [{ToDosTable}] (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT, TaskName TEXT, TargetUsers TEXT, DueDate TEXT, 
                        AdvanceDays INTEGER, Message TEXT, IsActive INTEGER DEFAULT 1);";

                    using (var cmd = new SQLiteCommand(sql1, conn)) cmd.ExecuteNonQuery();
                    using (var cmd = new SQLiteCommand(sql2, conn)) cmd.ExecuteNonQuery();
                    using (var cmd = new SQLiteCommand(sql3, conn)) cmd.ExecuteNonQuery();

                    // 動態升級表結構：新增循環任務欄位
                    var cols = new List<string>();
                    using (var cmd = new SQLiteCommand($"PRAGMA table_info([{ToDosTable}])", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) cols.Add(reader["name"].ToString());
                    }

                    if (!cols.Contains("IsRecurring")) {
                        using (var cmd = new SQLiteCommand($"ALTER TABLE [{ToDosTable}] ADD COLUMN IsRecurring INTEGER DEFAULT 0;", conn)) cmd.ExecuteNonQuery();
                        using (var cmd = new SQLiteCommand($"ALTER TABLE [{ToDosTable}] ADD COLUMN RecurType TEXT;", conn)) cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        public static DataTable GetAllRules()
        {
            DataTable dt = new DataTable();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"SELECT * FROM {RulesTable} ORDER BY IsActive DESC, Id ASC", conn))
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            } catch { }
            return dt;
        }

        public static DataTable GetAllToDos()
        {
            DataTable dt = new DataTable();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"SELECT * FROM {ToDosTable} ORDER BY IsActive DESC, DueDate ASC", conn))
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            } catch { }
            return dt;
        }

        // ==============================================================
        // 核心檢查與觸發機制 (整合 Rules 與 ToDos)
        // ==============================================================
        public static void CheckAndShowReminders()
        {
            InitDatabase();

            Task.Run(() => 
            {
                try
                {
                    string currentUser = Environment.UserName.Trim();
                    string todayStr = DateTime.Today.ToString("yyyy-MM-dd");
                    List<TriggeredReminder> triggeredList = new List<TriggeredReminder>();

                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();

                        // ---------------------------------------------------------
                        // 1. 處理資料表掃描規則 (RulesTable)
                        // ---------------------------------------------------------
                        DataTable activeRules = new DataTable();
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {RulesTable} WHERE IsActive = 1", conn))
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(activeRules);

                        foreach (DataRow rule in activeRules.Rows)
                        {
                            string targets = rule["TargetUsers"].ToString().Trim();
                            bool isTargetUser = CheckIsTargetUser(targets, currentUser);
                            if (!isTargetUser) continue;

                            int ruleId = Convert.ToInt32(rule["Id"]);
                            string dbName = rule["DbName"].ToString();
                            string tbName = rule["TableName"].ToString();
                            string dateCol = rule["DateCol"].ToString();
                            int advanceDays = Convert.ToInt32(rule["AdvanceDays"]);
                            string template = rule["MessageTemplate"].ToString();

                            DataTable sourceData = null;
                            try { sourceData = DataManager.GetTableData(dbName, tbName, "", "", ""); } catch { continue; }

                            if (sourceData == null || sourceData.Rows.Count == 0 || !sourceData.Columns.Contains(dateCol) || !sourceData.Columns.Contains("Id")) continue;

                            Dictionary<int, string> snoozeLogs = GetSnoozeLogs(conn, currentUser, ruleId);
                            Regex fieldRegex = new Regex(@"\[(.*?)\]");

                            foreach (DataRow srcRow in sourceData.Rows)
                            {
                                if (srcRow.RowState == DataRowState.Deleted) continue;

                                int recordId = Convert.ToInt32(srcRow["Id"]);
                                if (snoozeLogs.ContainsKey(recordId)) {
                                    if (string.Compare(snoozeLogs[recordId], todayStr) > 0) continue; 
                                }

                                string rawDate = srcRow[dateCol]?.ToString() ?? "";
                                DateTime? parsedDate = ParseUniversalDate(rawDate);
                                
                                if (parsedDate.HasValue)
                                {
                                    int daysDiff = (int)(parsedDate.Value.Date - DateTime.Today).TotalDays;
                                    if (daysDiff <= advanceDays)
                                    {
                                        string finalMsg = template;
                                        var matches = fieldRegex.Matches(template);
                                        foreach (Match m in matches) {
                                            string colName = m.Groups[1].Value;
                                            if (sourceData.Columns.Contains(colName)) {
                                                finalMsg = finalMsg.Replace(m.Value, srcRow[colName]?.ToString() ?? "");
                                            }
                                        }

                                        triggeredList.Add(new TriggeredReminder {
                                            RuleId = ruleId,
                                            RecordId = recordId,
                                            RuleName = rule["RuleName"].ToString(),
                                            Message = finalMsg,
                                            DaysLeft = daysDiff
                                        });
                                    }
                                }
                            }
                        }

                        // ---------------------------------------------------------
                        // 2. 處理自訂待辦事項 (CustomToDos) - 統一使用 RuleId = 0
                        // ---------------------------------------------------------
                        DataTable activeToDos = new DataTable();
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {ToDosTable} WHERE IsActive = 1", conn))
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(activeToDos);

                        Dictionary<int, string> todoSnoozeLogs = GetSnoozeLogs(conn, currentUser, 0); 

                        foreach (DataRow todo in activeToDos.Rows)
                        {
                            string targets = todo["TargetUsers"].ToString().Trim();
                            if (!CheckIsTargetUser(targets, currentUser)) continue;

                            int todoId = Convert.ToInt32(todo["Id"]);
                            if (todoSnoozeLogs.ContainsKey(todoId)) {
                                if (string.Compare(todoSnoozeLogs[todoId], todayStr) > 0) continue; 
                            }

                            string rawDate = todo["DueDate"]?.ToString() ?? "";
                            DateTime? parsedDate = ParseUniversalDate(rawDate);
                            int advanceDays = Convert.ToInt32(todo["AdvanceDays"]);
                            
                            bool isRecurring = todo.Table.Columns.Contains("IsRecurring") && todo["IsRecurring"] != DBNull.Value && Convert.ToInt32(todo["IsRecurring"]) == 1;
                            string recurType = todo.Table.Columns.Contains("RecurType") ? todo["RecurType"].ToString() : "";

                            if (parsedDate.HasValue)
                            {
                                DateTime targetDate = parsedDate.Value.Date;

                                // 循環任務邏輯：自動計算下一個應觸發的日期
                                if (isRecurring && !string.IsNullOrEmpty(recurType))
                                {
                                    targetDate = CalculateNextRecurringDate(targetDate, recurType);
                                }

                                int daysDiff = (int)(targetDate - DateTime.Today).TotalDays;
                                
                                if (daysDiff <= advanceDays)
                                {
                                    string recurLabel = isRecurring ? $"[循環任務:{recurType}] " : "";
                                    triggeredList.Add(new TriggeredReminder {
                                        RuleId = 0,               
                                        RecordId = todoId,        
                                        RuleName = $"{recurLabel}[待辦] {todo["TaskName"]}",
                                        Message = todo["Message"].ToString(),
                                        DaysLeft = daysDiff
                                    });
                                }
                            }
                        }
                    } 

                    // 3. 若有觸發提醒，喚起 UI
                    if (triggeredList.Count > 0)
                    {
                        if (Application.OpenForms.Count > 0)
                        {
                            Form mainForm = Application.OpenForms[0];
                            if (mainForm.InvokeRequired) {
                                mainForm.Invoke(new Action(() => ShowPopupUI(triggeredList, currentUser)));
                            } else {
                                ShowPopupUI(triggeredList, currentUser);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Reminder Engine Error: " + ex.Message);
                }
            });
        }

        private static DateTime CalculateNextRecurringDate(DateTime baseDate, string recurType)
        {
            DateTime today = DateTime.Today;
            DateTime nextDate = baseDate;

            if (nextDate < today)
            {
                if (recurType == "每天") {
                    nextDate = today;
                }
                else if (recurType == "每週") {
                    while (nextDate < today) nextDate = nextDate.AddDays(7);
                }
                else if (recurType == "每月") {
                    while (nextDate < today) nextDate = nextDate.AddMonths(1);
                }
                else if (recurType == "每年") {
                    while (nextDate < today) nextDate = nextDate.AddYears(1);
                }
            }
            return nextDate;
        }

        private static bool CheckIsTargetUser(string targets, string currentUser)
        {
            if (targets.Equals("ALL", StringComparison.OrdinalIgnoreCase)) return true;
            var users = targets.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(u => u.Trim()).ToList();
            return users.Contains(currentUser, StringComparer.OrdinalIgnoreCase);
        }

        private static Dictionary<int, string> GetSnoozeLogs(SQLiteConnection conn, string userName, int ruleId)
        {
            Dictionary<int, string> logs = new Dictionary<int, string>();
            using (var cmd = new SQLiteCommand($"SELECT RecordId, NextRemindDate FROM {LogsTable} WHERE UserName=@U AND RuleId=@R", conn)) {
                cmd.Parameters.AddWithValue("@U", userName);
                cmd.Parameters.AddWithValue("@R", ruleId);
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        logs[Convert.ToInt32(reader["RecordId"])] = reader["NextRemindDate"].ToString();
                    }
                }
            }
            return logs;
        }

        // ==============================================================
        // 🚀 終極修復版：使用 Master TableLayoutPanel 確保無重疊
        // ==============================================================
        private static void ShowPopupUI(List<TriggeredReminder> reminders, string userName)
        {
            foreach (Form openForm in Application.OpenForms)
            {
                if (openForm.Text == "🔔 系統智能提醒") return;
            }

            using (Form f = new Form())
            {
                f.Text = "🔔 系統智能提醒";
                f.Size = new Size(800, 650);
                f.StartPosition = FormStartPosition.CenterScreen;
                f.FormBorderStyle = FormBorderStyle.Sizable; 
                f.MaximizeBox = false;
                f.MinimizeBox = false;
                f.BackColor = Color.WhiteSmoke;
                f.TopMost = true; 

                // 🟢 建立全視窗的 Master 網格佈局 (TableLayoutPanel)
                // 這能 100% 保證三層架構絕對不會疊加在一起
                TableLayoutPanel masterTlp = new TableLayoutPanel {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 3,
                    Margin = new Padding(0)
                };
                masterTlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                
                // 第一層：標題 (自動適應文字高度)
                masterTlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                // 第二層：滾動清單區 (佔用剩下所有空間 100%)
                masterTlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                // 第三層：儲存按鈕區 (絕對高度 70px)
                masterTlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));

                // 1. 建立頂部標題區塊
                Label lblTop = new Label { 
                    Text = $"您共有 {reminders.Count} 筆待處理的系統提醒：", 
                    Dock = DockStyle.Fill, // 填滿網格第一層
                    Padding = new Padding(15, 15, 15, 15), 
                    Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                    ForeColor = Color.DarkRed,
                    BackColor = Color.WhiteSmoke
                };

                // 2. 建立中央滾動清單區塊
                FlowLayoutPanel flp = new FlowLayoutPanel { 
                    Dock = DockStyle.Fill, // 填滿網格第二層
                    AutoScroll = true, 
                    Padding = new Padding(15), 
                    FlowDirection = FlowDirection.TopDown, 
                    WrapContents = false,
                    BackColor = Color.WhiteSmoke
                };

                // 動態寬度適應
                flp.Resize += (s, e) => {
                    foreach (Control c in flp.Controls) {
                        if (c is Panel card) {
                            card.Width = flp.ClientSize.Width - flp.Padding.Left - flp.Padding.Right - card.Margin.Left - card.Margin.Right;
                        }
                    }
                };
                
                Dictionary<TriggeredReminder, ComboBox> actionMap = new Dictionary<TriggeredReminder, ComboBox>();

                foreach (var rm in reminders.OrderBy(r => r.DaysLeft))
                {
                    // 🟢 卡片主體
                    Panel cardPanel = new Panel { 
                        Width = 720, 
                        AutoSize = true, 
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        BackColor = Color.White, 
                        Margin = new Padding(0, 0, 0, 15),
                        Padding = new Padding(15, 15, 15, 20) 
                    };
                    cardPanel.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, cardPanel.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

                    // 卡片內部網格
                    TableLayoutPanel tlpCard = new TableLayoutPanel {
                        Dock = DockStyle.Top, 
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        ColumnCount = 1,
                        RowCount = 3,
                        Margin = new Padding(0)
                    };
                    tlpCard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                    tlpCard.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    tlpCard.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    tlpCard.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); 

                    // ============== 第一列：標籤與規則名稱 ==============
                    FlowLayoutPanel flpHeader = new FlowLayoutPanel {
                        AutoSize = true,
                        WrapContents = false,
                        Margin = new Padding(0, 0, 0, 10)
                    };
                    
                    string statusTag = rm.DaysLeft < 0 ? $"[已逾期 {Math.Abs(rm.DaysLeft)} 天]" : (rm.DaysLeft == 0 ? "[今日到期]" : $"[還有 {rm.DaysLeft} 天]");
                    Color tagColor = rm.DaysLeft < 0 ? Color.Crimson : (rm.DaysLeft <= 7 ? Color.DarkOrange : Color.DarkSlateBlue);

                    Label lblTag = new Label { Text = statusTag, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = tagColor, AutoSize = true };
                    Label lblRule = new Label { Text = $"標籤：{rm.RuleName}", Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(10, 2, 0, 0) };
                    
                    flpHeader.Controls.Add(lblTag);
                    flpHeader.Controls.Add(lblRule);
                    tlpCard.Controls.Add(flpHeader, 0, 0);

                    // ============== 第二列：訊息內容 ==============
                    Label lblMsg = new Label { 
                        Text = rm.Message, 
                        Font = new Font("Microsoft JhengHei UI", 12F), 
                        AutoSize = true, 
                        Dock = DockStyle.Fill, 
                        Margin = new Padding(5, 0, 0, 15) 
                    };
                    tlpCard.Controls.Add(lblMsg, 0, 1);

                    // ============== 第三列：下拉選單 ==============
                    ComboBox cboAction = new ComboBox { 
                        Width = 260, 
                        DropDownStyle = ComboBoxStyle.DropDownList, 
                        Font = new Font("Microsoft JhengHei UI", 11F), 
                        Anchor = AnchorStyles.Right, // 鎖死在右側
                        Margin = new Padding(0, 0, 5, 0)
                    };
                    cboAction.Items.AddRange(new string[] { "本次忽略 (下次開啟再提醒)", "今天不再提醒 (延至明天)", "3 天後再提醒", "7 天後再提醒", "本訊息永久不再提醒" });
                    cboAction.SelectedIndex = 0;
                    actionMap[rm] = cboAction;
                    
                    tlpCard.Controls.Add(cboAction, 0, 2);

                    // 組裝
                    cardPanel.Controls.Add(tlpCard);
                    flp.Controls.Add(cardPanel);
                }

                // 3. 建立底部面板 (儲存按鈕)
                Panel pnlBottom = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20, 10, 20, 10), BackColor = Color.White };
                Button btnSave = new Button { Text = "💾 儲存設定並關閉視窗", Dock = DockStyle.Fill, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                pnlBottom.Controls.Add(btnSave);

                // 🟢 依序放入 Master 網格，完全避免重疊
                masterTlp.Controls.Add(lblTop, 0, 0);     // 放在第 0 列
                masterTlp.Controls.Add(flp, 0, 1);        // 放在第 1 列
                masterTlp.Controls.Add(pnlBottom, 0, 2);  // 放在第 2 列

                f.Controls.Add(masterTlp);

                // 🟢 防止焦點亂跳
                f.Shown += (s, e) => {
                    btnSave.Focus(); 
                };

                // 儲存邏輯
                btnSave.Click += (s, e) => {
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var trans = conn.BeginTransaction()) {
                                foreach (var kvp in actionMap) {
                                    int actionIdx = kvp.Value.SelectedIndex;
                                    if (actionIdx == 0) continue; 
                                    
                                    string nextDate = "";
                                    if (actionIdx == 1) nextDate = DateTime.Today.AddDays(1).ToString("yyyy-MM-dd");
                                    else if (actionIdx == 2) nextDate = DateTime.Today.AddDays(3).ToString("yyyy-MM-dd");
                                    else if (actionIdx == 3) nextDate = DateTime.Today.AddDays(7).ToString("yyyy-MM-dd");
                                    else if (actionIdx == 4) nextDate = "2099-12-31"; 

                                    using (var cmdDel = new SQLiteCommand("DELETE FROM UserReminderLogs WHERE UserName=@U AND RuleId=@R AND RecordId=@Rec", conn, trans)) {
                                        cmdDel.Parameters.AddWithValue("@U", userName);
                                        cmdDel.Parameters.AddWithValue("@R", kvp.Key.RuleId);
                                        cmdDel.Parameters.AddWithValue("@Rec", kvp.Key.RecordId);
                                        cmdDel.ExecuteNonQuery();
                                    }

                                    using (var cmdIns = new SQLiteCommand("INSERT INTO UserReminderLogs (UserName, RuleId, RecordId, NextRemindDate) VALUES (@U, @R, @Rec, @ND)", conn, trans)) {
                                        cmdIns.Parameters.AddWithValue("@U", userName);
                                        cmdIns.Parameters.AddWithValue("@R", kvp.Key.RuleId);
                                        cmdIns.Parameters.AddWithValue("@Rec", kvp.Key.RecordId);
                                        cmdIns.Parameters.AddWithValue("@ND", nextDate);
                                        cmdIns.ExecuteNonQuery();
                                    }
                                }
                                trans.Commit();
                            }
                        }
                        f.DialogResult = DialogResult.OK;
                    } catch (Exception ex) {
                        MessageBox.Show("儲存延遲提醒設定失敗：" + ex.Message);
                    }
                };

                f.ShowDialog();
            }
        }

        private static DateTime? ParseUniversalDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return null;
            dateStr = dateStr.Trim().Replace("/", "-");
            Regex twRegex = new Regex(@"^(?<year>\d{2,3})-(?<month>\d{1,2})-(?<day>\d{1,2})(?:\s+.*)?$");
            Match matchTw = twRegex.Match(dateStr);
            if (matchTw.Success) {
                if (int.TryParse(matchTw.Groups["year"].Value, out int twYear)) {
                    if (twYear < 200) {
                        int westernYear = twYear + 1911;
                        dateStr = $"{westernYear}-{matchTw.Groups["month"].Value.PadLeft(2, '0')}-{matchTw.Groups["day"].Value.PadLeft(2, '0')}"; 
                    }
                }
            }
            if (DateTime.TryParse(dateStr, out DateTime result)) return result;
            return null;
        }
    }
}
