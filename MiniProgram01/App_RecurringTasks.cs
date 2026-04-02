using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Drawing.Printing;

public class App_RecurringTasks : UserControl {
    private MainForm parentForm;
    private App_TodoList todoApp;
    private string recurringFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_recurring.txt");
    private FlowLayoutPanel taskPanel;
    private Timer checkTimer;

    public string digestType { get; set; } = "不提醒";
    public string digestTimeStr { get; set; } = "08:00";
    public string lastDigestDate { get; set; } = "";
    public int advanceDays { get; set; } = 0;
    public string scanFrequency { get; set; } = "10分鐘";

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public class RecurringTask { 
        public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate, Note, TaskType; 
    }
    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; 
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 4 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label() { Text = "週期任務", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5,0,0,0) };
        Button btnViewAll = new Button() { Text = "全部檢視", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke };
        btnViewAll.Click += (s, e) => OpenAllTasksView();
        Button btnAdd = new Button() { Text = "新增任務", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand, BackColor = Color.White };
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this).ShowDialog(); };
        Button btnSet = new Button() { Text = "設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        header.Controls.AddRange(new Control[] { lblTitle, btnViewAll, btnAdd, btnSet });
        this.Controls.Add(header);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - 25;
            if (safeWidth > 0) foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
        };
        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadTasks();
        checkTimer = new Timer() { Interval = GetTimerInterval(scanFrequency), Enabled = true };
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks();
    }

    private int GetTimerInterval(string freq) {
        switch (freq) {
            case "即時": return 1000; case "1分鐘": return 60000; case "5分鐘": return 300000;
            case "10分鐘": return 600000; case "1小時": return 3600000; case "12小時": return 43200000;
            case "1天": return 86400000; default: return 600000;
        }
    }

    public void OpenAllTasksView() { this.Invoke(new Action(() => { new AllTasksViewWindow(this).Show(); })); }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 25 : 450;
        foreach (var t in tasks) {
            Panel card = new Panel() { Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 45), Margin = new Padding(5, 5, 5, 8), BackColor = Color.FromArgb(248, 248, 250) };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, AutoSize = true, Padding = new Padding(5) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button btnEdit = new Button() { Text = "調", Dock = DockStyle.Top, Height = 28, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnEdit.Click += (s, e) => { int idx = tasks.IndexOf(t); if(idx != -1) { new EditRecurringTaskWindow(this, idx, t).ShowDialog(); RefreshUI(); } };
            Button btnDel = new Button() { Text = "✕", Dock = DockStyle.Top, Height = 28, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnDel.Click += (s, e) => { if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) DeleteTask(t); };
            Button btnNote = new Button() { Text = "註", Dock = DockStyle.Top, Height = 28, FlatStyle = FlatStyle.Flat };
            btnNote.BackColor = string.IsNullOrEmpty(t.Note) ? Color.FromArgb(230, 230, 230) : Color.FromArgb(255, 193, 7);
            btnNote.Click += (s, e) => { string n = ShowNoteEditBox(t.Name, t.Note); if (n != null) { t.Note = n; SaveTasks(); RefreshUI(); } };

            string typeTag = $"[{t.TaskType}] ";
            Label lbl = new Label() { Text = typeTag + string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Font = MainFont };
            
            tlp.Controls.Add(btnEdit, 0, 0); tlp.Controls.Add(btnDel, 1, 0); tlp.Controls.Add(btnNote, 2, 0); tlp.Controls.Add(lbl, 3, 0);
            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }

    public void AddNewTask(string name, string month, string date, string time, string note, string type) {
        tasks.Add(new RecurringTask() { Name = name, MonthStr = month, DateStr = date, TimeStr = time, Note = note, TaskType = type, LastTriggeredDate = "" });
        SaveTasks(); RefreshUI();
    }

    public void UpdateTask(int index, string name, string month, string date, string time, string note, string type) {
        if (index >= 0 && index < tasks.Count) {
            tasks[index].Name = name; tasks[index].MonthStr = month; tasks[index].DateStr = date; 
            tasks[index].TimeStr = time; tasks[index].Note = note; tasks[index].TaskType = type;
            SaveTasks(); RefreshUI();
        }
    }

    public void DeleteTask(RecurringTask task) { if (tasks.Contains(task)) { tasks.Remove(task); SaveTasks(); RefreshUI(); } }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays, string sFreq) {
        digestType = dType; digestTimeStr = dTime; advanceDays = aDays; scanFrequency = sFreq;
        checkTimer.Enabled = false; checkTimer.Interval = GetTimerInterval(sFreq); checkTimer.Enabled = true;
        SaveTasks(); MessageBox.Show("設定儲存成功！");
    }

    private void CheckTasks() {
        DateTime now = DateTime.Now; bool needsSave = false;
        List<RecurringTask> toRemove = new List<RecurringTask>();

        foreach (var t in tasks) {
            DateTime target;
            if (TryGetNextTriggerTime(t, now, out target)) {
                DateTime triggerThreshold = target.AddDays(-advanceDays);
                if (now >= triggerThreshold) {
                    string targetDateStr = target.ToString("yyyy-MM-dd");
                    if (t.LastTriggeredDate != targetDateStr) {
                        string prefix = advanceDays > 0 ? string.Format("[預排-{0}] ", target.ToString("MM/dd")) : "";
                        todoApp.AddTask(prefix + t.Name, "Black", "週期觸發", t.Note); 
                        t.LastTriggeredDate = targetDateStr; needsSave = true;
                        parentForm.AlertTab(1);
                        
                        // 如果是單次或到期日，觸發後從清單中移除
                        if (t.TaskType == "單次" || t.TaskType == "到期日") toRemove.Add(t);
                    }
                }
            }
        }

        if (toRemove.Count > 0) {
            foreach (var r in toRemove) tasks.Remove(r);
            needsSave = true;
            this.Invoke(new Action(() => RefreshUI()));
        }

        if (digestType != "不提醒" && DateTime.TryParseExact(digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtDigest)) {
            DateTime targetDigest = new DateTime(now.Year, now.Month, now.Day, dtDigest.Hour, dtDigest.Minute, 0);
            bool shouldTrigger = (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday && now >= targetDigest) || (digestType == "每月1號" && now.Day == 1 && now >= targetDigest);
            if (shouldTrigger && lastDigestDate != now.ToString("yyyy-MM-dd")) { lastDigestDate = now.ToString("yyyy-MM-dd"); needsSave = true; OpenAllTasksView(); }
        }
        if (needsSave) SaveTasks();
    }

    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target) {
        target = now;
        try {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]), m = int.Parse(timeParts[1]);
            if (t.MonthStr == "每天") { target = new DateTime(now.Year, now.Month, now.Day, h, m, 0); if (now > target) target = target.AddDays(1); return true; }
            if (t.MonthStr == "每週") {
                var dow = new Dictionary<string, DayOfWeek>() {{"一",DayOfWeek.Monday},{"二",DayOfWeek.Tuesday},{"三",DayOfWeek.Wednesday},{"四",DayOfWeek.Thursday},{"五",DayOfWeek.Friday},{"六",DayOfWeek.Saturday},{"日",DayOfWeek.Sunday}};
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                while (target.DayOfWeek != dow[t.DateStr] || now > target) target = target.AddDays(1);
                return true;
            }
            if (t.MonthStr == "每月" || t.MonthStr.EndsWith("月")) {
                int month = t.MonthStr == "每月" ? now.Month : int.Parse(t.MonthStr.Replace("月",""));
                int day = (t.DateStr == "月底") ? DateTime.DaysInMonth(now.Year, month) : int.Parse(t.DateStr);
                target = new DateTime(now.Year, month, Math.Min(day, DateTime.DaysInMonth(now.Year, month)), h, m, 0);
                if (now > target) target = t.MonthStr == "每月" ? target.AddMonths(1) : target.AddYears(1);
                return true;
            }
        } catch { } return false;
    }

    public void SaveTasks() {
        List<string> lines = new List<string>(){ string.Format("#DIGEST|{0}|{1}|{2}|{3}|{4}", digestType, digestTimeStr, lastDigestDate, advanceDays, scanFrequency) };
        foreach(var t in tasks) lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate, EncodeBase64(t.Note), t.TaskType));
        File.WriteAllLines(recurringFile, lines);
    }

    private void LoadTasks() {
        if(!File.Exists(recurringFile)) return;
        tasks.Clear();
        foreach(var l in File.ReadAllLines(recurringFile)) {
            var p = l.Split('|');
            if(l.StartsWith("#DIGEST")) { 
                digestType = p[1]; digestTimeStr = p[2]; lastDigestDate = p[3]; 
                if(p.Length >= 5) advanceDays = int.Parse(p[4]);
                if(p.Length >= 6) scanFrequency = p[5];
            } else if(p.Length >= 4) {
                tasks.Add(new RecurringTask() { 
                    Name = p[0], MonthStr = p[1], DateStr = p[2], TimeStr = p[3], 
                    LastTriggeredDate = p.Length > 4 ? p[4] : "", 
                    Note = p.Length > 5 ? DecodeBase64(p[5]) : "",
                    TaskType = p.Length > 6 ? p[6] : "循環" 
                });
            }
        }
        RefreshUI();
    }

    private string ShowNoteEditBox(string name, string current) {
        Form f = new Form() { Width = 400, Height = 350, Text = "編輯備註", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog };
        TextBox txt = new TextBox() { Left = 15, Top = 50, Width = 350, Height = 180, Multiline = true, Text = current };
        Button btn = new Button() { Text = "儲存", Left = 265, Top = 250, Width = 100, Height = 35, DialogResult =
