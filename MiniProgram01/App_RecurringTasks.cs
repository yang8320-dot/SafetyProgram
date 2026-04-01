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
    
    // 【新增】掃描頻率設定變數，預設為 10分鐘
    public string scanFrequency { get; set; } = "10分鐘";

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public class RecurringTask { 
        public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate, Note; 
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
        
        // 【修改】將「排程設定」文字精簡為「設定」
        Button btnSet = new Button() { Text = "設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        header.Controls.AddRange(new Control[] { lblTitle, btnViewAll, btnAdd, btnSet });
        this.Controls.Add(header);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - 25;
            if (safeWidth > 0) {
                foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
            }
        };

        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadTasks();
        
        // 【新增】初始化時載入對應的頻率間隔
        checkTimer = new Timer() { Interval = GetTimerInterval(scanFrequency), Enabled = true };
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks();
    }

    // 【新增】負責將文字轉換為毫秒的對應器
    private int GetTimerInterval(string freq) {
        switch (freq) {
            case "即時": return 1000;       // 1秒
            case "1分鐘": return 60000;     // 60秒
            case "5分鐘": return 300000;
            case "10分鐘": return 600000;
            case "1小時": return 3600000;
            case "12小時": return 43200000;
            case "1天": return 86400000;
            default: return 600000;       // 預設 10分鐘
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

            Button btnEdit = new Button() { Text = "調", Dock = DockStyle.Top, Height = 28, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0,0,3,0) };
            btnEdit.FlatAppearance.BorderSize = 0; 
            btnEdit.Click += (s, e) => { 
                int currentIdx = tasks.IndexOf(t);
                if(currentIdx != -1) { new EditRecurringTaskWindow(this, currentIdx, t).ShowDialog(); RefreshUI(); }
            };

            Button btnDel = new Button() { Text = "✕", Dock = DockStyle.Top, Height = 28, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0,0,3,0) };
            btnDel.FlatAppearance.BorderSize = 0; 
            btnDel.Click += (s, e) => { if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) DeleteTask(t); };

            Button btnNote = new Button() { Text = "註", Dock = DockStyle.Top, Height = 28, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0,0,3,0), Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold) };
            btnNote.FlatAppearance.BorderSize = 0;
            
            Action updateNoteStyle = () => {
                if (!string.IsNullOrEmpty(t.Note)) { btnNote.BackColor = Color.FromArgb(255, 193, 7); btnNote.ForeColor = Color.Black; }
                else { btnNote.BackColor = Color.FromArgb(230, 230, 230); btnNote.ForeColor = Color.Gray; }
            };
            updateNoteStyle();
            btnNote.Click += (s, e) => {
                string newNote = ShowNoteEditBox(t.Name, t.Note);
                if (newNote != null) { t.Note = newNote; updateNoteStyle(); SaveTasks(); }
            };

            Label lbl = new Label() { Text = string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Font = MainFont };
            
            tlp.Controls.Add(btnEdit, 0, 0); 
            tlp.Controls.Add(btnDel, 1, 0); 
            tlp.Controls.Add(btnNote, 2, 0);
            tlp.Controls.Add(lbl, 3, 0);
            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }

    public void AddNewTask(string name, string month, string date, string time, string note) {
        tasks.Add(new RecurringTask() { Name = name, MonthStr = month, DateStr = date, TimeStr = time, Note = note, LastTriggeredDate = "" });
        SaveTasks(); RefreshUI();
    }

    public void UpdateTask(int index, string name, string month, string date, string time, string note) {
        if (index >= 0 && index < tasks.Count) {
            tasks[index].Name = name; tasks[index].MonthStr = month; tasks[index].DateStr = date; tasks[index].TimeStr = time; tasks[index].Note = note;
            SaveTasks(); RefreshUI();
        }
    }

    public void DeleteTask(RecurringTask task) { if (tasks.Contains(task)) { tasks.Remove(task); SaveTasks(); RefreshUI(); } }

    // 【修改】接收新的掃描頻率，並立刻套用更新 Timer
    public void UpdateGlobalSettings(string dType, string dTime, int aDays, string sFreq) {
        digestType = dType; digestTimeStr = dTime; advanceDays = aDays; scanFrequency = sFreq;
        
        // 變更頻率後立刻重啟 Timer 套用新頻率
        checkTimer.Enabled = false;
        checkTimer.Interval = GetTimerInterval(sFreq);
        checkTimer.Enabled = true;
        
        SaveTasks(); MessageBox.Show("設定儲存成功！");
    }

    private void CheckTasks() {
        DateTime now = DateTime.Now; bool needsSave = false;
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
                    }
                }
            }
        }
        if (digestType != "不提醒") {
            DateTime dtDigest;
            if (DateTime.TryParseExact(digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtDigest)) {
                DateTime targetDigest = new DateTime(now.Year, now.Month, now.Day, dtDigest.Hour, dtDigest.Minute, 0);
                bool shouldTrigger = (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday && now >= targetDigest) || (digestType == "每月1號" && now.Day == 1 && now >= targetDigest);
                if (shouldTrigger && lastDigestDate != now.ToString("yyyy-MM-dd")) { lastDigestDate = now.ToString("yyyy-MM-dd"); needsSave = true; OpenAllTasksView(); }
            }
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
                if (!dow.ContainsKey(t.DateStr)) return false;
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

    // 【修改】將 scanFrequency 也寫入設定檔儲存 (加入在第 5 個位置)
    private void SaveTasks() {
        List<string> lines = new List<string>(){ string.Format("#DIGEST|{0}|{1}|{2}|{3}|{4}", digestType, digestTimeStr, lastDigestDate, advanceDays, scanFrequency) };
        foreach(var t in tasks) lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}|{5}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate, EncodeBase64(t.Note)));
        File.WriteAllLines(recurringFile, lines);
    }

    // 【修改】讀取設定檔時，把 scanFrequency 給抓出來
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
                tasks.Add(new RecurringTask() { Name = p[0], MonthStr = p[1], DateStr = p[2], TimeStr = p[3], LastTriggeredDate = p.Length > 4 ? p[4] : "", Note = p.Length > 5 ? DecodeBase64(p[5]) : "" });
            }
        }
        RefreshUI();
    }

    private string ShowNoteEditBox(string name, string current) {
        Form f = new Form() { Width = 400, Height = 350, Text = "編輯週期備註", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog };
        TextBox txt = new TextBox() { Left = 15, Top = 50, Width = 350, Height = 180, Multiline = true, WordWrap = true, ScrollBars = ScrollBars.Vertical, Font = MainFont, Text = current };
        Button btn = new Button() { Text = "儲存", Left = 265, Top = 250, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        f.Controls.AddRange(new Control[] { new Label(){Text="【"+name+"】", Left=15, Top=15, AutoSize=true, Font=new Font(MainFont, FontStyle.Bold)}, txt, btn });
        return f.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }

    private string EncodeBase64(string t) => string.IsNullOrEmpty(t) ? "" : Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(t));
    private string DecodeBase64(string b) { try { return string.IsNullOrEmpty(b) ? "" : System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(b)); } catch { return ""; } }
}

// ==========================================
// 檢視全部任務
// ==========================================
public class AllTasksViewWindow : Form {
    private App_RecurringTasks parentControl;
    private FlowLayoutPanel flow;
    public AllTasksViewWindow(App_RecurringTasks parent) {
        this.parentControl = parent; this.Text = "週期任務總覽"; this.Width = 820; this.Height = 800; this.BackColor = Color.White;
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 60, BackColor = Color.WhiteSmoke, ColumnCount = 2 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170f));
        Label lbl = new Label() { Text = "週期任務排程總覽", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(15,0,0,0), Font = new Font("Microsoft JhengHei UI", 16f, FontStyle.Bold) };
        Button btnPrint = new Button() { Text = "轉存 PDF / 列印", Width = 150, Height = 35, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Right, Margin = new Padding(0,0,15,0) };
        header.Controls.Add(lbl, 0, 0); header.Controls.Add(btnPrint, 1, 0); this.Controls.Add(header);
        flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15, 20, 15, 15), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        flow.Resize += (s, e) => { int w = flow.ClientSize.Width - 35; if (w > 0) foreach (Control c in flow.Controls) if (c is GroupBox gb) gb.Width = w; };
        this.Controls.Add(flow); flow.BringToFront(); RefreshData();
    }
    public void RefreshData() {
        flow.Controls.Clear(); var tasks = parentControl.tasks;
        AddGroup("每天觸發", tasks.Where(t => t.MonthStr == "每天").ToList());
        AddGroup("每週觸發", tasks.Where(t => t.MonthStr == "每週").ToList());
        AddGroup("每月觸發", tasks.Where(t => t.MonthStr == "每月").ToList());
        for (int i = 1; i <= 12; i++) AddGroup(i + "月 限定", tasks.Where(t => t.MonthStr == i + "月").ToList());
    }
    private void AddGroup(string header, List<App_RecurringTasks.RecurringTask> sub) {
        if (sub.Count == 0) return;
        GroupBox gb = new GroupBox() { Text = "【 " + header + " 】", Font = new Font("Microsoft JhengHei UI", 12f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 122, 255), AutoSize = true, Width = flow.ClientSize.Width - 35, Margin = new Padding(10, 10, 10, 25), Padding = new Padding(15, 25, 15, 15) };
        FlowLayoutPanel inner = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true };
        foreach (var t in sub) {
            TableLayoutPanel row = new TableLayoutPanel() { Width = gb.Width - 40, AutoSize = true, ColumnCount = 4, Margin = new Padding(0,0,0,8) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            Button bE = new Button() { Text = "調", Height = 28, Dock = DockStyle.Top, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            bE.Click += (s, e) => { if(parentControl.tasks.IndexOf(t) != -1) { new EditRecurringTaskWindow(parentControl, parentControl.tasks.IndexOf(t), t).ShowDialog(); RefreshData(); } };
            Button bD = new Button() { Text = "✕", Height = 28, Dock = DockStyle.Top, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            bD.Click += (s, e) => { if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { parentControl.DeleteTask(t); RefreshData(); } };
            Button bN = new Button() { Text = "註", Height = 28, Dock = DockStyle.Top, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold) };
            if (!string.IsNullOrEmpty(t.Note)) { bN.BackColor = Color.FromArgb(255,193,7); bN.ForeColor = Color.Black; } else { bN.BackColor = Color.FromArgb(230,230,230); bN.ForeColor = Color.Gray; }
            bN.Click += (s, e) => { 
                Form nf = new Form() { Width = 400, Height = 350, Text = "任務備註", StartPosition = FormStartPosition.CenterScreen };
                TextBox nt = new TextBox() { Left = 15, Top = 50, Width = 350, Height = 180, Multiline = true, Text = t.Note };
                Button nb = new Button() { Text = "儲存", Left = 265, Top = 250, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0,122,255), ForeColor = Color.White };
                nf.Controls.AddRange(new Control[] { new Label(){Text="【"+t.Name+"】", Left=15, Top=15, AutoSize=true}, nt, nb });
                if(nf.ShowDialog() == DialogResult.OK) { t.Note = nt.Text; parentControl.SaveTasks(); RefreshData(); }
            };
            row.Controls.Add(bE, 0, 0); row.Controls.Add(bD, 1, 0); row.Controls.Add(bN, 2, 0);
            row.Controls.Add(new Label() { Text = string.Format("[{0}] {1}  {2}", t.TimeStr, t.DateStr, t.Name), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Padding = new Padding(0,8,0,8) }, 3, 0);
            inner.Controls.Add(row);
        }
        gb.Controls.Add(inner); container.Controls.Add(gb);
    }
}

// ==========================================
// 視窗：新增/調整
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD;
    private DateTimePicker dtp;
    public AddRecurringTaskWindow(App_RecurringTasks p) {
        this.parent = p; this.Text = "新增週期任務"; this.Width = 360; this.Height = 500; this.StartPosition = FormStartPosition.CenterScreen;
        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };
        f.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true });
        txtN = new TextBox() { Width = 290 }; f.Controls.Add(txtN);
        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", AutoSize = true, Margin = new Padding(0,10,0,0) });
        txtNote = new TextBox() { Width = 290, Height = 80, Multiline = true, ScrollBars = ScrollBars.Vertical }; f.Controls.Add(txtNote);
        f.Controls.Add(new Label() { Text = "週期類型：", AutoSize = true, Margin = new Padding(0,10,0,0) });
        cmM = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.AddRange(new string[]{"每天","每週","每月"}); for(int i=1;i<=12;i++) cmM.Items.Add(i+"月");
        f.Controls.Add(cmM);
        f.Controls.Add(new Label() { Text = "日期條件：", AutoSize = true, Margin = new Padding(0,10,0,0) });
        cmD = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList }; f.Controls.Add(cmD);
        cmM.SelectedIndexChanged += (s,e) => {
            cmD.Items.Clear();
            if(cmM.Text=="每天") { cmD.Items.Add("每日"); cmD.Enabled=false; }
            else if(cmM.Text=="每週") { cmD.Items.AddRange(new string[]{"一","二","三","四","五","六","日"}); cmD.Enabled=true; }
            else { for(int i=1;i<=31;i++) cmD.Items.Add(i.ToString()); cmD.Items.Add("月底"); cmD.Enabled=true; }
            cmD.SelectedIndex = 0;
        }; cmM.SelectedIndex = 0;
        dtp = new DateTimePicker() { Width = 290, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Value = DateTime.Today.AddHours(9), Margin = new Padding(0,15,0,0) };
        f.Controls.Add(dtp);
        Button btn = new Button() { Text = "建立任務", Width = 290, Height = 40, BackColor = Color.FromArgb(0,122,255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(0,20,0,0) };
        btn.Click += (s,e) => { if(!string.IsNullOrWhiteSpace(txtN.Text)) { parent.AddNewTask(txtN.Text, cmM.Text, cmD.Text, dtp.Value.ToString("HH:mm"), txtNote.Text); this.Close(); } };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}

public class EditRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private int idx;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD;
    private DateTimePicker dtp;
    public EditRecurringTaskWindow(App_RecurringTasks p, int i, App_RecurringTasks.RecurringTask t) {
        this.parent = p; this.idx = i; this.Text = "調整任務條件"; this.Width = 360; this.Height = 500; this.StartPosition = FormStartPosition.CenterScreen;
        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };
        txtN = new TextBox() { Width = 290, Text = t.Name }; f.Controls.Add(new Label(){Text="任務名稱："}); f.Controls.Add(txtN);
        txtNote = new TextBox() { Width = 290, Height = 80, Multiline = true, Text = t.Note }; f.Controls.Add(new Label(){Text="詳細說明 (註)：", Margin=new Padding(0,10,0,0)}); f.Controls.Add(txtNote);
        cmM = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.AddRange(new string[]{"每天","每週","每月"}); for(int k=1;k<=12;k++) cmM.Items.Add(k+"月");
        cmM.Text = t.MonthStr; f.Controls.Add(new Label(){Text="週期類型：", Margin=new Padding(0,10,0,0)}); f.Controls.Add(cmM);
        cmD = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList }; f.Controls.Add(new Label(){Text="日期條件：", Margin=new Padding(0,10,0,0)}); f.Controls.Add(cmD);
        cmM.SelectedIndexChanged += (s,e) => {
            cmD.Items.Clear();
            if(cmM.Text=="每天") { cmD.Items.Add("每日"); }
            else if(cmM.Text=="每週") { cmD.Items.AddRange(new string[]{"一","二","三","四","五","六","日"}); }
            else { for(int k=1;k<=31;k++) cmD.Items.Add(k.ToString()); cmD.Items.Add("月底"); }
            if(cmD.Items.Count>0) cmD.SelectedIndex=0;
        }; cmD.Text = t.DateStr;
        dtp = new DateTimePicker() { Width = 290, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        DateTime dtv; if(DateTime.TryParseExact(t.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtv)) dtp.Value = dtv;
        f.Controls.Add(dtp);
        Button btn = new Button() { Text = "儲存修改", Width = 290, Height = 40, BackColor = Color.FromArgb(0,153,76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(0,20,0,0) };
        btn.Click += (s,e) => { parent.UpdateTask(idx, txtN.Text, cmM.Text, cmD.Text, dtp.Value.ToString("HH:mm"), txtNote.Text); this.Close(); };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}

// 【修改】將視窗加高，並新增分隔線與掃描頻率選項
public class RecurringSettingsWindow : Form {
    private App_RecurringTasks parent;
    private ComboBox cmDig, cmAdv, cmScan;
    private DateTimePicker dtp;
    
    public RecurringSettingsWindow(App_RecurringTasks p) {
        this.parent = p; this.Text = "全域排程設定"; this.Width = 350; 
        this.Height = 330; // 加高視窗以容納新選項
        this.StartPosition = FormStartPosition.CenterScreen;
        
        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };
        
        FlowLayoutPanel r1 = new FlowLayoutPanel() { AutoSize = true };
        r1.Controls.Add(new Label() { Text = "所有任務提前", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmAdv = new ComboBox() { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList }; for (int i = 0; i <= 7; i++) cmAdv.Items.Add(i.ToString());
        cmAdv.Text = p.advanceDays.ToString(); r1.Controls.Add(cmAdv); r1.Controls.Add(new Label() { Text = "天加入待辦", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        f.Controls.Add(r1);
        
        FlowLayoutPanel r2 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 10, 0, 10) };
        r2.Controls.Add(new Label() { Text = "視窗摘要提醒：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmDig = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList }; cmDig.Items.AddRange(new string[]{"不提醒", "每週一", "每月1號"}); cmDig.Text = p.digestType;
        dtp = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        DateTime dtv; if(DateTime.TryParseExact(p.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtv)) dtp.Value = dtv;
        r2.Controls.AddRange(new Control[]{ cmDig, dtp }); f.Controls.Add(r2);
        
        // 【新增】視覺分隔線
        Label sep = new Label() { AutoSize = false, Height = 2, Width = 290, BorderStyle = BorderStyle.Fixed3D, Margin = new Padding(0, 5, 0, 15) };
        f.Controls.Add(sep);

        // 【新增】掃描頻率選項
        FlowLayoutPanel r3 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
        r3.Controls.Add(new Label() { Text = "待辦事項掃描頻率：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmScan = new ComboBox() { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
        cmScan.Items.AddRange(new string[] { "即時", "1分鐘", "5分鐘", "10分鐘", "1小時", "12小時", "1天" });
        cmScan.Text = p.scanFrequency; // 讀取目前的設定
        r3.Controls.Add(cmScan);
        f.Controls.Add(r3);

        Button btn = new Button() { Text = "儲存所有設定", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btn.Click += (s, e) => { 
            // 儲存時將掃描頻率一併傳回
            p.UpdateGlobalSettings(cmDig.Text, dtp.Value.ToString("HH:mm"), int.Parse(cmAdv.Text), cmScan.Text); 
            this.Close(); 
        };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}
