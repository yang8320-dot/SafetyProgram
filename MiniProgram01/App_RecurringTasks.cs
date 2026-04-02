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

            // 顯示類型標籤
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
                        
                        // 【邏輯修正】：如果是「單次」或已達「到期日」，觸發後標記移除
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
        Button btn = new Button() { Text = "儲存", Left = 265, Top = 250, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        f.Controls.AddRange(new Control[] { new Label(){Text="【"+name+"】", Left=15, Top=15, AutoSize=true, Font=new Font(MainFont, FontStyle.Bold)}, txt, btn });
        return f.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }
    private string EncodeBase64(string t) => string.IsNullOrEmpty(t) ? "" : Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(t));
    private string DecodeBase64(string b) { try { return string.IsNullOrEmpty(b) ? "" : System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(b)); } catch { return ""; } }
}

// ==========================================
// 視窗：新增/調整 (加入功能選擇選單)
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD, cmType;
    private DateTimePicker dtp;
    public AddRecurringTaskWindow(App_RecurringTasks p) {
        this.parent = p; this.Text = "新增任務"; this.Width = 360; this.Height = 520; this.StartPosition = FormStartPosition.CenterScreen;
        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };
        f.Controls.Add(new Label() { Text = "任務名稱：" }); txtN = new TextBox() { Width = 290 }; f.Controls.Add(txtN);
        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, 10, 0, 0) });
        txtNote = new TextBox() { Width = 290, Height = 60, Multiline = true }; f.Controls.Add(txtNote);
        
        // 【新增功能選單】
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, 10, 0, 0) });
        cmType = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
        cmType.Items.AddRange(new string[] { "循環", "單次", "到期日" }); cmType.SelectedIndex = 0; f.Controls.Add(cmType);

        f.Controls.Add(new Label() { Text = "週期類型：", Margin = new Padding(0, 10, 0, 0) });
        cmM = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.AddRange(new string[]{"每天","每週","每月"}); for(int i=1;i<=12;i++) cmM.Items.Add(i+"月");
        f.Controls.Add(cmM); cmD = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList }; f.Controls.Add(cmD);
        cmM.SelectedIndexChanged += (s,e) => {
            cmD.Items.Clear();
            if(cmM.Text=="每天") { cmD.Items.Add("每日"); cmD.Enabled=false; }
            else if(cmM.Text=="每週") { cmD.Items.AddRange(new string[]{"一","二","三","四","五","六","日"}); cmD.Enabled=true; }
            else { for(int i=1;i<=31;i++) cmD.Items.Add(i.ToString()); cmD.Items.Add("月底"); cmD.Enabled=true; }
            cmD.SelectedIndex = 0;
        }; cmM.SelectedIndex = 0;
        dtp = new DateTimePicker() { Width = 290, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Value = DateTime.Today.AddHours(9), Margin = new Padding(0,10,0,0) };
        f.Controls.Add(dtp);
        Button btn = new Button() { Text = "建立任務", Width = 290, Height = 40, BackColor = Color.FromArgb(0,122,255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(0,20,0,0) };
        btn.Click += (s,e) => { if(!string.IsNullOrWhiteSpace(txtN.Text)) { parent.AddNewTask(txtN.Text, cmM.Text, cmD.Text, dtp.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); this.Close(); } };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}

public class EditRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private int idx;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD, cmType;
    private DateTimePicker dtp;
    public EditRecurringTaskWindow(App_RecurringTasks p, int i, App_RecurringTasks.RecurringTask t) {
        this.parent = p; this.idx = i; this.Text = "調整任務"; this.Width = 360; this.Height = 520; this.StartPosition = FormStartPosition.CenterScreen;
        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };
        f.Controls.Add(new Label() { Text = "任務名稱：" }); txtN = new TextBox() { Width = 290, Text = t.Name }; f.Controls.Add(txtN);
        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, 10, 0, 0) });
        txtNote = new TextBox() { Width = 290, Height = 60, Multiline = true, Text = t.Note }; f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, 10, 0, 0) });
        cmType = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
        cmType.Items.AddRange(new string[] { "循環", "單次", "到期日" }); cmType.Text = t.TaskType; f.Controls.Add(cmType);

        f.Controls.Add(new Label() { Text = "週期類型：" });
        cmM = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.AddRange(new string[]{"每天","每週","每月"}); for(int k=1;k<=12;k++) cmM.Items.Add(k+"月");
        cmM.Text = t.MonthStr; f.Controls.Add(cmM);
        cmD = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList }; f.Controls.Add(cmD);
        cmM.SelectedIndexChanged += (s,e) => {
            cmD.Items.Clear();
            if(cmM.Text=="每天") { cmD.Items.Add("每日"); }
            else if(cmM.Text=="每週") { cmD.Items.AddRange(new string[]{"一","二","三","四","五","六","日"}); }
            else { for(int k=1;k<=31;k++) cmD.Items.Add(k.ToString()); cmD.Items.Add("月底"); }
            if(cmD.Items.Count>0) cmD.SelectedIndex=0;
        }; cmD.Text = t.DateStr;
        dtp = new DateTimePicker() { Width = 290, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        if(DateTime.TryParseExact(t.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtp.Value = dtv;
        f.Controls.Add(dtp);
        Button btn = new Button() { Text = "儲存修改", Width = 290, Height = 40, BackColor = Color.FromArgb(0,153,76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 20, 0, 0) };
        btn.Click += (s,e) => { parent.UpdateTask(idx, txtN.Text, cmM.Text, cmD.Text, dtp.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); this.Close(); };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}
// (AllTasksViewWindow 和 RecurringSettingsWindow 部分保持原樣即可，不需更動)
