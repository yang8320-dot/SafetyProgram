/// FILE: Safety_System/ReminderEngine.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
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
                    
                    using (var cmd = new SQLiteCommand(sql1, conn)) cmd.ExecuteNonQuery();
                    using (var cmd = new SQLiteCommand(sql2, conn)) cmd.ExecuteNonQuery();
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

        // ==============================================================
        // 核心檢查與觸發機制
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

                    // 1. 取得當前使用者需要套用的啟用的規則
                    DataTable activeRules = new DataTable();
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand($"SELECT * FROM {RulesTable} WHERE IsActive = 1", conn))
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(activeRules);
                    }

                    List<TriggeredReminder> triggeredList = new List<TriggeredReminder>();

                    foreach (DataRow rule in activeRules.Rows)
                    {
                        string targets = rule["TargetUsers"].ToString().Trim();
                        bool isTargetUser = false;
                        
                        if (targets.Equals("ALL", StringComparison.OrdinalIgnoreCase)) isTargetUser = true;
                        else {
                            var users = targets.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(u => u.Trim()).ToList();
                            if (users.Contains(currentUser, StringComparer.OrdinalIgnoreCase)) isTargetUser = true;
                        }

                        if (!isTargetUser) continue;

                        int ruleId = Convert.ToInt32(rule["Id"]);
                        string dbName = rule["DbName"].ToString();
                        string tbName = rule["TableName"].ToString();
                        string dateCol = rule["DateCol"].ToString();
                        int advanceDays = Convert.ToInt32(rule["AdvanceDays"]);
                        string template = rule["MessageTemplate"].ToString();

                        // 2. 去對應資料表撈取資料
                        DataTable sourceData = null;
                        try {
                            // 不傳入時間，撈取全表 (可以優化為透過 SQL 過濾，但考慮到 SQLite 日期格式混亂，在記憶體過濾最安全)
                            sourceData = DataManager.GetTableData(dbName, tbName, "", "", "");
                        } catch { continue; }

                        if (sourceData == null || sourceData.Rows.Count == 0 || !sourceData.Columns.Contains(dateCol) || !sourceData.Columns.Contains("Id")) continue;

                        // 3. 抓取使用者的延遲提醒紀錄
                        Dictionary<int, string> snoozeLogs = new Dictionary<int, string>();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand($"SELECT RecordId, NextRemindDate FROM {LogsTable} WHERE UserName=@U AND RuleId=@R", conn)) {
                                cmd.Parameters.AddWithValue("@U", currentUser);
                                cmd.Parameters.AddWithValue("@R", ruleId);
                                using (var reader = cmd.ExecuteReader()) {
                                    while (reader.Read()) {
                                        snoozeLogs[Convert.ToInt32(reader["RecordId"])] = reader["NextRemindDate"].ToString();
                                    }
                                }
                            }
                        }

                        // 4. 比對與樣板生成
                        Regex fieldRegex = new Regex(@"\[(.*?)\]");

                        foreach (DataRow srcRow in sourceData.Rows)
                        {
                            if (srcRow.RowState == DataRowState.Deleted) continue;

                            int recordId = Convert.ToInt32(srcRow["Id"]);
                            
                            // 檢查是否被 Snooze (延遲)
                            if (snoozeLogs.ContainsKey(recordId)) {
                                string nextDateStr = snoozeLogs[recordId];
                                if (string.Compare(nextDateStr, todayStr) > 0) continue; // 還沒到下次提醒時間
                            }

                            string rawDate = srcRow[dateCol]?.ToString() ?? "";
                            DateTime? parsedDate = ParseUniversalDate(rawDate);
                            
                            if (parsedDate.HasValue)
                            {
                                int daysDiff = (int)(parsedDate.Value.Date - DateTime.Today).TotalDays;
                                
                                // 只要天數小於或等於設定的提前天數，就視為觸發 (過期的負數也會觸發)
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

                    // 5. 若有觸發提醒，喚起 UI
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

        private static void ShowPopupUI(List<TriggeredReminder> reminders, string userName)
        {
            using (Form f = new Form())
            {
                f.Text = "🔔 系統智能提醒";
                f.Size = new Size(800, 600);
                f.StartPosition = FormStartPosition.CenterScreen;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;
                f.BackColor = Color.WhiteSmoke;
                f.TopMost = true; // 強制顯示在最上層

                Label lblTop = new Label { Text = $"您共有 {reminders.Count} 筆待處理的系統提醒：", Dock = DockStyle.Top, Padding = new Padding(15), Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkRed, Height = 60 };
                f.Controls.Add(lblTop);

                FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false };
                
                // 儲存每個項目使用者選擇的處理動作
                Dictionary<TriggeredReminder, ComboBox> actionMap = new Dictionary<TriggeredReminder, ComboBox>();

                foreach (var rm in reminders.OrderBy(r => r.DaysLeft))
                {
                    Panel pnl = new Panel { Width = 740, AutoSize = true, MinimumSize = new Size(740, 80), BackColor = Color.White, Margin = new Padding(0, 0, 0, 10), Padding = new Padding(10) };
                    pnl.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnl.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

                    string statusTag = rm.DaysLeft < 0 ? $"[已逾期 {Math.Abs(rm.DaysLeft)} 天]" : (rm.DaysLeft == 0 ? "[今日到期]" : $"[還有 {rm.DaysLeft} 天]");
                    Color tagColor = rm.DaysLeft < 0 ? Color.Crimson : (rm.DaysLeft <= 7 ? Color.DarkOrange : Color.DarkSlateBlue);

                    Label lblTag = new Label { Text = statusTag, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = tagColor, Location = new Point(10, 15), AutoSize = true };
                    Label lblRule = new Label { Text = $"標籤：{rm.RuleName}", Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.DimGray, Location = new Point(150, 18), AutoSize = true };
                    
                    Label lblMsg = new Label { Text = rm.Message, Font = new Font("Microsoft JhengHei UI", 12F), Location = new Point(10, 45), MaximumSize = new Size(500, 0), AutoSize = true };
                    
                    ComboBox cboAction = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(530, 40) };
                    cboAction.Items.AddRange(new string[] { "本次忽略 (下次開啟再提醒)", "今天不再提醒 (延至明天)", "3 天後再提醒", "7 天後再提醒", "本訊息永久不再提醒" });
                    cboAction.SelectedIndex = 0;
                    
                    actionMap[rm] = cboAction;

                    pnl.Controls.Add(lblTag);
                    pnl.Controls.Add(lblRule);
                    pnl.Controls.Add(lblMsg);
                    pnl.Controls.Add(cboAction);
                    flp.Controls.Add(pnl);
                }

                f.Controls.Add(flp);

                Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 70, Padding = new Padding(20, 10, 20, 10), BackColor = Color.White };
                Button btnSave = new Button { Text = "💾 儲存設定並關閉視窗", Dock = DockStyle.Fill, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                
                btnSave.Click += (s, e) => {
                    try {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var trans = conn.BeginTransaction()) {
                                foreach (var kvp in actionMap) {
                                    int actionIdx = kvp.Value.SelectedIndex;
                                    if (actionIdx == 0) continue; // 本次忽略，不寫入資料庫
                                    
                                    string nextDate = "";
                                    if (actionIdx == 1) nextDate = DateTime.Today.AddDays(1).ToString("yyyy-MM-dd");
                                    else if (actionIdx == 2) nextDate = DateTime.Today.AddDays(3).ToString("yyyy-MM-dd");
                                    else if (actionIdx == 3) nextDate = DateTime.Today.AddDays(7).ToString("yyyy-MM-dd");
                                    else if (actionIdx == 4) nextDate = "2099-12-31"; // 永久不提醒

                                    string sql = @"INSERT INTO UserReminderLogs (UserName, RuleId, RecordId, NextRemindDate) 
                                                   VALUES (@U, @R, @Rec, @ND) 
                                                   ON CONFLICT(UserName, RuleId, RecordId) DO UPDATE SET NextRemindDate=@ND";
                                    
                                    // 確保有 Unique Index 支持 UPSERT，若無則先刪再塞
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

                pnlBottom.Controls.Add(btnSave);
                f.Controls.Add(pnlBottom);

                f.ShowDialog();
            }
        }

        // 萬用日期解析器
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
