/*
 * 檔案功能：週期任務管理與自動推播模組 (SQLite 升級版)
 * 對應選單名稱：週期任務
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表名稱：RecurringTasks (排程清單), GlobalSettings (全域設定)
 */

using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_RecurringTasks : UserControl
{
    // --- 核心變數與參考 ---
    private MainForm parentForm;
    private dynamic todoApp; 
    private FlowLayoutPanel taskPanel;
    private System.Windows.Forms.Timer checkTimer;

    // --- 全域排程設定 (預設值) ---
    public string digestType = "不提醒";
    public string digestTimeStr = "08:00";
    public string lastDigestDate = "";
    public int advanceDays = 0;
    public string scanFrequency = "10分鐘";

    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleOrange = Color.FromArgb(255, 149, 0);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);
    private static Font SmallFont = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Regular);

    // --- 資料模型 ---
    public class RecurringTask
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string MonthStr { get; set; } 
        public string DateStr { get; set; }  
        public string TimeStr { get; set; }  
        public string LastTriggeredDate { get; set; }
        public string Note { get; set; }
        public string TaskType { get; set; } 
    }

    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, dynamic todoApp)
    {
        this.parentForm = mainForm;
        this.todoApp = todoApp;

        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15);

        InitializeUI();
        _ = LoadTasksAndSettingsAsync();
    }

    private void InitializeUI()
    {
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 3, BackColor = Color.Transparent };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));

        Label lblTitle = new Label() { Text = "週期排程任務", Font = new Font(MainFont.FontFamily, 14f, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5, 0, 0, 0) };

        Button btnSet = new Button() { Text = "全域設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = Color.White, ForeColor = AppleBlue, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold), Margin = new Padding(0, 5, 10, 5) };
        btnSet.FlatAppearance.BorderSize = 1; btnSet.FlatAppearance.BorderColor = AppleBlue;
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        Button btnAdd = new Button() { Text = "新增", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold), Margin = new Padding(0, 5, 0, 5) };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += (s, e) => { new AddEditRecurringTaskWindow(this).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0); header.Controls.Add(btnSet, 1, 0); header.Controls.Add(btnAdd, 2, 0);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = AppleBgColor };
        taskPanel.Resize += (s, e) => { int safeWidth = taskPanel.ClientSize.Width - 20; if (safeWidth > 0) { foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth; } };

        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent });
        this.Controls.Add(header);
        taskPanel.BringToFront();
    }

    public void RefreshUI()
    {
        if (this.InvokeRequired) { this.Invoke(new Action(RefreshUI)); return; }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var t in tasks)
        {
            Panel card = new Panel() { Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 65), Margin = new Padding(0, 0, 0, 15), BackColor = Color.White, Padding = new Padding(10) };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 2, AutoSize = true };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));

            Label lblTitle = new Label() { Text = $"[{t.TaskType}] {t.Name}", Font = BoldFont, AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 0, 0, 5) };
            tlp.Controls.Add(lblTitle, 0, 0);

            string scheduleStr = t.MonthStr == "特定日期" ? $"{t.DateStr} {t.TimeStr}" : $"{t.MonthStr} {t.DateStr} {t.TimeStr}";
            Label lblDetails = new Label() { Text = $"觸發條件: {scheduleStr} (上次觸發: {t.LastTriggeredDate})", Font = SmallFont, ForeColor = Color.DarkGray, AutoSize = true, Dock = DockStyle.Fill };
            tlp.Controls.Add(lblDetails, 0, 1);

            Button btnNote = new Button() { Text = "註", Dock = DockStyle.Fill, BackColor = string.IsNullOrWhiteSpace(t.Note) ? Color.FromArgb(230, 230, 230) : AppleOrange, ForeColor = string.IsNullOrWhiteSpace(t.Note) ? Color.Gray : Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(5) };
            btnNote.FlatAppearance.BorderSize = 0; tlp.SetRowSpan(btnNote, 2);
            btnNote.Click += (s, e) => { MessageBox.Show(string.IsNullOrWhiteSpace(t.Note) ? "無備註" : t.Note, $"備註: {t.Name}", MessageBoxButtons.OK, MessageBoxIcon.Information); };
            tlp.Controls.Add(btnNote, 1, 0);

            Button btnEdit = new Button() { Text = "編輯", Dock = DockStyle.Fill, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(5) };
            btnEdit.FlatAppearance.BorderSize = 0; tlp.SetRowSpan(btnEdit, 2);
            btnEdit.Click += (s, e) => { new AddEditRecurringTaskWindow(this, t).ShowDialog(); };
            tlp.Controls.Add(btnEdit, 2, 0);

            Button btnDel = new Button() { Text = "刪除", Dock = DockStyle.Fill, BackColor = AppleRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(5, 5, 0, 5) };
            btnDel.FlatAppearance.BorderSize = 0; tlp.SetRowSpan(btnDel, 2);
            btnDel.Click += async (s, e) =>
            {
                if (MessageBox.Show("確定移除此任務？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    await DeleteTaskAsync(t.Id);
                    tasks.Remove(t);
                    RefreshUI();
                }
            };
            tlp.Controls.Add(btnDel, 3, 0);

            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }

    // ==========================================
    // SQLite 資料存取邏輯
    // ==========================================
    private async Task LoadTasksAndSettingsAsync()
    {
        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    // 1. 載入全域設定
                    using (var cmd = new SQLiteCommand("SELECT SettingKey, SettingValue FROM GlobalSettings", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string key = reader.GetString(0); string val = reader.GetString(1);
                            if (key == "DigestType") digestType = val;
                            if (key == "DigestTimeStr") digestTimeStr = val;
                            if (key == "LastDigestDate") lastDigestDate = val;
                            if (key == "AdvanceDays" && int.TryParse(val, out int aDays)) advanceDays = aDays;
                            if (key == "ScanFrequency") scanFrequency = val;
                        }
                    }

                    // 2. 載入排程任務
                    var tempList = new List<RecurringTask>();
                    using (var cmd = new SQLiteCommand("SELECT Id, Name, MonthStr, DateStr, TimeStr, LastTriggeredDate, Note, TaskType FROM RecurringTasks", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tempList.Add(new RecurringTask
                            {
                                Id = reader.GetString(0), Name = reader.GetString(1), MonthStr = reader.GetString(2),
                                DateStr = reader.GetString(3), TimeStr = reader.GetString(4), LastTriggeredDate = reader.GetString(5),
                                Note = reader.GetString(6), TaskType = reader.GetString(7)
                            });
                        }
                    }
                    tasks = tempList;
                }
            });

            RefreshUI();

            checkTimer = new System.Windows.Forms.Timer();
            UpdateTimerFrequency();
            checkTimer.Tick += (s, e) => CheckTasks();
            checkTimer.Start();
            CheckTasks();
        }
        catch (Exception ex) { MessageBox.Show($"載入週期任務失敗: {ex.Message}"); }
    }

    public async Task SaveTaskAsync(RecurringTask t)
    {
        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand(@"INSERT OR REPLACE INTO RecurringTasks (Id, Name, MonthStr, DateStr, TimeStr, LastTriggeredDate, Note, TaskType) 
                                                         VALUES (@Id, @Name, @MonthStr, @DateStr, @TimeStr, @LastTriggeredDate, @Note, @TaskType)", conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", t.Id); cmd.Parameters.AddWithValue("@Name", t.Name);
                        cmd.Parameters.AddWithValue("@MonthStr", t.MonthStr); cmd.Parameters.AddWithValue("@DateStr", t.DateStr);
                        cmd.Parameters.AddWithValue("@TimeStr", t.TimeStr); cmd.Parameters.AddWithValue("@LastTriggeredDate", t.LastTriggeredDate);
                        cmd.Parameters.AddWithValue("@Note", t.Note ?? ""); cmd.Parameters.AddWithValue("@TaskType", t.TaskType);
                        cmd.ExecuteNonQuery();
                    }
                }
            });
        }
        catch (Exception ex) { MessageBox.Show($"儲存任務失敗: {ex.Message}"); }
    }

    public async Task DeleteTaskAsync(string id)
    {
        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("DELETE FROM RecurringTasks WHERE Id = @Id", conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", id);
                        cmd.ExecuteNonQuery();
                    }
                }
            });
        }
        catch (Exception ex) { MessageBox.Show($"刪除任務失敗: {ex.Message}"); }
    }

    public async Task SaveSettingsAsync()
    {
        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        void Set(string k, string v)
                        {
                            using (var cmd = new SQLiteCommand("INSERT OR REPLACE INTO GlobalSettings (SettingKey, SettingValue) VALUES (@k, @v)", conn, tx))
                            {
                                cmd.Parameters.AddWithValue("@k", k); cmd.Parameters.AddWithValue("@v", v); cmd.ExecuteNonQuery();
                            }
                        }
                        Set("DigestType", digestType); Set("DigestTimeStr", digestTimeStr);
                        Set("LastDigestDate", lastDigestDate); Set("AdvanceDays", advanceDays.ToString());
                        Set("ScanFrequency", scanFrequency);
                        tx.Commit();
                    }
                }
            });
        }
        catch (Exception ex) { MessageBox.Show($"儲存設定失敗: {ex.Message}"); }
    }

    // ==========================================
    // 背景排程偵測引擎 (非同步與執行緒安全)
    // ==========================================
    public void UpdateTimerFrequency()
    {
        if (checkTimer != null)
        {
            switch (scanFrequency)
            {
                case "即時": checkTimer.Interval = 1000; break;
                case "1分鐘": checkTimer.Interval = 60000; break;
                case "5分鐘": checkTimer.Interval = 300000; break;
                case "10分鐘": checkTimer.Interval = 600000; break;
                case "1小時": checkTimer.Interval = 3600000; break;
                default: checkTimer.Interval = 600000; break;
            }
        }
    }

    private async void CheckTasks()
    {
        DateTime now = DateTime.Now;
        bool needsRefresh = false;
        List<RecurringTask> toRemove = new List<RecurringTask>();

        foreach (var t in tasks)
        {
            if (TryGetNextTriggerTime(t, now, out DateTime target))
            {
                DateTime triggerThreshold = target.AddDays(-advanceDays);
                if (now >= triggerThreshold)
                {
                    string targetDateStr = target.ToString("yyyy-MM-dd");
                    if (t.LastTriggeredDate != targetDateStr)
                    {
                        string prefix = advanceDays > 0 ? $"[預排-{target:MM/dd}] " : "";
                        
                        if (todoApp != null)
                        {
                            if (parentForm.InvokeRequired) parentForm.Invoke(new Action(async () => await todoApp.ReceiveTaskAsync(prefix + t.Name, now.ToString("yyyy-MM-dd HH:mm:ss"))));
                            else await todoApp.ReceiveTaskAsync(prefix + t.Name, now.ToString("yyyy-MM-dd HH:mm:ss"));
                        }

                        t.LastTriggeredDate = targetDateStr;
                        await SaveTaskAsync(t); // 更新資料庫的觸發日期
                        needsRefresh = true;

                        if (t.TaskType == "單次") toRemove.Add(t);
                    }
                }
            }
        }

        if (toRemove.Count > 0)
        {
            foreach (var r in toRemove) { await DeleteTaskAsync(r.Id); tasks.Remove(r); }
            needsRefresh = true;
        }

        if (needsRefresh) RefreshUI();
    }

    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target)
    {
        target = DateTime.MinValue;
        try
        {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]); int m = int.Parse(timeParts[1]);

            if (t.MonthStr == "每天")
            {
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                if (now > target) target = target.AddDays(1);
                return true;
            }
            else if (t.MonthStr == "每週")
            {
                int targetDayOfWeek = "日一二三四五六".IndexOf(t.DateStr);
                int daysToAdd = targetDayOfWeek - (int)now.DayOfWeek;
                if (daysToAdd < 0 || (daysToAdd == 0 && now.TimeOfDay.TotalMinutes > h * 60 + m)) daysToAdd += 7;
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0).AddDays(daysToAdd);
                return true;
            }
            else if (t.MonthStr == "每月")
            {
                int validDay = t.DateStr == "月底" ? DateTime.DaysInMonth(now.Year, now.Month) : Math.Min(int.Parse(t.DateStr), DateTime.DaysInMonth(now.Year, now.Month));
                target = new DateTime(now.Year, now.Month, validDay, h, m, 0);
                if (now > target)
                {
                    DateTime nextMonth = now.AddMonths(1);
                    validDay = t.DateStr == "月底" ? DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month) : Math.Min(int.Parse(t.DateStr), DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month));
                    target = new DateTime(nextMonth.Year, nextMonth.Month, validDay, h, m, 0);
                }
                return true;
            }
            else if (t.MonthStr == "特定日期")
            {
                DateTime specDate = DateTime.Parse(t.DateStr);
                target = new DateTime(specDate.Year, specDate.Month, specDate.Day, h, m, 0);
                return true;
            }
        }
        catch { }
        return false;
    }
}

// ==========================================
// 視窗：新增/編輯任務 (iOS 風格整合版)
// ==========================================
public class AddEditRecurringTaskWindow : Form
{
    private App_RecurringTasks parent;
    private App_RecurringTasks.RecurringTask editTarget;
    
    private TextBox txtName, txtNote;
    private ComboBox cmbType, cmbCycleType, cmbDate;
    private DateTimePicker dtpSpecificDate, dtpTime;

    public AddEditRecurringTaskWindow(App_RecurringTasks parent, App_RecurringTasks.RecurringTask item = null)
    {
        this.parent = parent; this.editTarget = item;

        this.Text = item == null ? "新增週期任務" : "編輯週期任務";
        this.Width = 450; this.Height = 580; this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247); this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20), AutoScroll = true, WrapContents = false };
        void AddLabel(string text) => flow.Controls.Add(new Label() { Text = text, AutoSize = true, Margin = new Padding(0, 10, 0, 5) });

        AddLabel("任務名稱：");
        txtName = new TextBox() { Width = 380, BorderStyle = BorderStyle.FixedSingle, Text = item?.Name ?? "" }; flow.Controls.Add(txtName);

        AddLabel("執行模式：");
        cmbType = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 380 };
        cmbType.Items.AddRange(new string[] { "循環", "單次" }); cmbType.SelectedItem = item?.TaskType ?? "循環"; flow.Controls.Add(cmbType);

        AddLabel("循環頻率 / 特定日期：");
        cmbCycleType = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 380 };
        cmbCycleType.Items.AddRange(new string[] { "每天", "每週", "每月", "特定日期" }); cmbCycleType.SelectedItem = item?.MonthStr ?? "每天"; flow.Controls.Add(cmbCycleType);

        AddLabel("細節日期：");
        Panel datePanel = new Panel() { Width = 380, Height = 35, Margin = new Padding(0, 0, 0, 10) };
        cmbDate = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        dtpSpecificDate = new DateTimePicker() { Format = DateTimePickerFormat.Short, Dock = DockStyle.Fill, Visible = false };
        datePanel.Controls.Add(dtpSpecificDate); datePanel.Controls.Add(cmbDate); flow.Controls.Add(datePanel);

        AddLabel("觸發時間：");
        dtpTime = new DateTimePicker() { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 380 };
        if (item != null && DateTime.TryParseExact(item.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtpTime.Value = dtv;
        flow.Controls.Add(dtpTime);

        AddLabel("備註 (選填)：");
        txtNote = new TextBox() { Width = 380, Height = 80, Multiline = true, BorderStyle = BorderStyle.FixedSingle, Text = item?.Note ?? "" }; flow.Controls.Add(txtNote);

        cmbCycleType.SelectedIndexChanged += (s, e) =>
        {
            string sel = cmbCycleType.Text;
            if (sel == "特定日期") { cmbDate.Visible = false; dtpSpecificDate.Visible = true; }
            else
            {
                dtpSpecificDate.Visible = false; cmbDate.Visible = true; cmbDate.Items.Clear();
                if (sel == "每天") cmbDate.Items.Add("每日");
                else if (sel == "每週") cmbDate.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" });
                else if (sel == "每月") { for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i.ToString()); cmbDate.Items.Add("月底"); }
                if (cmbDate.Items.Count > 0) cmbDate.SelectedIndex = 0;
            }
        };
        cmbCycleType.SelectedIndex = cmbCycleType.Items.IndexOf(item?.MonthStr ?? "每天");
        if (item != null && item.MonthStr != "特定日期" && cmbDate.Items.Contains(item.DateStr)) cmbDate.SelectedItem = item.DateStr;
        if (item != null && item.MonthStr == "特定日期" && DateTime.TryParse(item.DateStr, out DateTime sDate)) dtpSpecificDate.Value = sDate;

        Button btnSave = new Button() { Text = "儲存任務", Width = 380, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), Margin = new Padding(0, 20, 0, 0) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入任務名稱！"); return; }
            string finalDateStr = cmbCycleType.Text == "特定日期" ? dtpSpecificDate.Value.ToString("yyyy-MM-dd") : cmbDate.Text;
            
            App_RecurringTasks.RecurringTask targetTask;
            if (editTarget == null)
            {
                targetTask = new App_RecurringTasks.RecurringTask()
                {
                    Id = Guid.NewGuid().ToString("N"), Name = txtName.Text.Trim(), MonthStr = cmbCycleType.Text,
                    DateStr = finalDateStr, TimeStr = dtpTime.Value.ToString("HH:mm"), LastTriggeredDate = "", Note = txtNote.Text, TaskType = cmbType.Text
                };
                parent.tasks.Add(targetTask);
            }
            else
            {
                targetTask = editTarget;
                targetTask.Name = txtName.Text.Trim(); targetTask.MonthStr = cmbCycleType.Text;
                targetTask.DateStr = finalDateStr; targetTask.TimeStr = dtpTime.Value.ToString("HH:mm");
                targetTask.Note = txtNote.Text; targetTask.TaskType = cmbType.Text;
            }

            await parent.SaveTaskAsync(targetTask);
            parent.RefreshUI();
            this.Close();
        };

        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}

// ==========================================
// 視窗：全域設定 (iOS 風格)
// ==========================================
public class RecurringSettingsWindow : Form
{
    private App_RecurringTasks parent;
    private NumericUpDown nudAdvance;
    private ComboBox cmbFreq;

    public RecurringSettingsWindow(App_RecurringTasks parent)
    {
        this.parent = parent;
        this.Text = "排程全域設定";
        this.Width = 350; this.Height = 280; this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247); this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };

        flow.Controls.Add(new Label() { Text = "提前幾天推播至待辦事項？", AutoSize = true, Margin = new Padding(0, 5, 0, 5) });
        nudAdvance = new NumericUpDown() { Width = 280, Minimum = 0, Maximum = 30, Value = parent.advanceDays }; flow.Controls.Add(nudAdvance);

        flow.Controls.Add(new Label() { Text = "背景偵測頻率：", AutoSize = true, Margin = new Padding(0, 15, 0, 5) });
        cmbFreq = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 280 };
        cmbFreq.Items.AddRange(new string[] { "即時", "1分鐘", "5分鐘", "10分鐘", "1小時" });
        cmbFreq.SelectedItem = parent.scanFrequency; flow.Controls.Add(cmbFreq);

        Button btnSave = new Button() { Text = "儲存設定", Width = 280, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), Margin = new Padding(0, 25, 0, 0) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            parent.advanceDays = (int)nudAdvance.Value;
            parent.scanFrequency = cmbFreq.Text;
            parent.UpdateTimerFrequency();
            await parent.SaveSettingsAsync();
            this.Close();
        };

        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}
