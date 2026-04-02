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

        TableLayoutPanel header = new TableLayoutPanel();
        header.Dock = DockStyle.Top;
        header.Height = 45;
        header.ColumnCount = 4;
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label();
        lblTitle.Text = "週期任務";
        lblTitle.Font = new Font(MainFont, FontStyle.Bold);
        lblTitle.Dock = DockStyle.Fill;
        lblTitle.TextAlign = ContentAlignment.MiddleLeft;
        lblTitle.Padding = new Padding(5, 0, 0, 0);

        Button btnViewAll = new Button();
        btnViewAll.Text = "全部檢視";
        btnViewAll.Dock = DockStyle.Fill;
        btnViewAll.FlatStyle = FlatStyle.Flat;
        btnViewAll.Margin = new Padding(2, 8, 2, 8);
        btnViewAll.Cursor = Cursors.Hand;
        btnViewAll.BackColor = Color.WhiteSmoke;
        // 【修正】改用 .Show() 釋放主視窗
        btnViewAll.Click += (s, e) => { new AllTasksViewWindow(this).Show(); }; 

        Button btnAdd = new Button();
        btnAdd.Text = "新增任務";
        btnAdd.Dock = DockStyle.Fill;
        btnAdd.FlatStyle = FlatStyle.Flat;
        btnAdd.Margin = new Padding(2, 8, 2, 8);
        btnAdd.Cursor = Cursors.Hand;
        btnAdd.BackColor = Color.White;
        // 【修正】改用 .Show() 釋放主視窗
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this).Show(); };

        Button btnSet = new Button();
        btnSet.Text = "設定";
        btnSet.Dock = DockStyle.Fill;
        btnSet.FlatStyle = FlatStyle.Flat;
        btnSet.BackColor = Color.Gainsboro;
        btnSet.Margin = new Padding(2, 8, 8, 8);
        btnSet.Cursor = Cursors.Hand;
        // 【修正】改用 .Show() 釋放主視窗
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).Show(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnViewAll, 1, 0);
        header.Controls.Add(btnAdd, 2, 0);
        header.Controls.Add(btnSet, 3, 0);
        this.Controls.Add(header);

        taskPanel = new FlowLayoutPanel();
        taskPanel.Dock = DockStyle.Fill;
        taskPanel.AutoScroll = true;
        taskPanel.FlowDirection = FlowDirection.TopDown;
        taskPanel.WrapContents = false;
        taskPanel.BackColor = Color.White;
        
        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - 25;
            if (safeWidth > 0) {
                foreach (Control c in taskPanel.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
            }
        };
        
        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadTasks();
        checkTimer = new Timer();
        checkTimer.Interval = GetTimerInterval(scanFrequency);
        checkTimer.Enabled = true;
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks();
    }

    private int GetTimerInterval(string freq) {
        switch (freq) {
            case "即時": return 1000; 
            case "1分鐘": return 60000; 
            case "5分鐘": return 300000;
            case "10分鐘": return 600000; 
            case "1小時": return 3600000; 
            case "12小時": return 43200000;
            case "1天": return 86400000; 
            default: return 600000;
        }
    }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 25 : 450;
        
        foreach (var t in tasks) {
            Panel card = new Panel();
            card.Width = startWidth;
            card.AutoSize = true;
            card.MinimumSize = new Size(0, 45);
            card.Margin = new Padding(5, 5, 5, 8);
            card.BackColor = Color.FromArgb(248, 248, 250);

            TableLayoutPanel tlp = new TableLayoutPanel();
            tlp.Dock = DockStyle.Fill;
            tlp.ColumnCount = 4;
            tlp.RowCount = 1;
            tlp.AutoSize = true;
            tlp.Padding = new Padding(5);
            
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button btnEdit = new Button();
            btnEdit.Text = "調";
            btnEdit.Dock = DockStyle.Top;
            btnEdit.Height = 28;
            btnEdit.BackColor = AppleBlue;
            btnEdit.ForeColor = Color.White;
            btnEdit.FlatStyle = FlatStyle.Flat;
            // 【修正】改用 .Show() 釋放主視窗 (RefreshUI已內建在存檔函式中)
            btnEdit.Click += (s, e) => { 
                int idx = tasks.IndexOf(t); 
                if(idx != -1) new EditRecurringTaskWindow(this, idx, t).Show(); 
            };

            Button btnDel = new Button();
            btnDel.Text = "✕";
            btnDel.Dock = DockStyle.Top;
            btnDel.Height = 28;
            btnDel.BackColor = Color.IndianRed;
            btnDel.ForeColor = Color.White;
            btnDel.FlatStyle = FlatStyle.Flat;
            btnDel.Click += (s, e) => { 
                if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) {
                    DeleteTask(t); 
                }
            };

            Button btnNote = new Button();
            btnNote.Text = "註";
            btnNote.Dock = DockStyle.Top;
            btnNote.Height = 28;
            btnNote.FlatStyle = FlatStyle.Flat;
            btnNote.BackColor = string.IsNullOrEmpty(t.Note) ? Color.FromArgb(230, 230, 230) : Color.FromArgb(255, 193, 7);
            btnNote.Click += (s, e) => { 
                string n = ShowNoteEditBox(t.Name, t.Note); 
                if (n != null) { 
                    t.Note = n; 
                    SaveTasks(); 
                    RefreshUI(); 
                } 
            };

            string typeTag = string.Format("[{0}] ", t.TaskType);
            // 判斷是否為單次/到期日的特定日期顯示方式
            string timeInfo = t.MonthStr == "特定日期" 
                ? string.Format("[{0} {1}]", t.DateStr, t.TimeStr) 
                : string.Format("[{0} {1} {2}]", t.MonthStr, t.DateStr, t.TimeStr);

            Label lbl = new Label();
            lbl.Text = typeTag + timeInfo + " " + t.Name;
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = ContentAlignment.MiddleLeft;
            lbl.AutoSize = true;
            lbl.Font = MainFont;
            
            tlp.Controls.Add(btnEdit, 0, 0); 
            tlp.Controls.Add(btnDel, 1, 0); 
            tlp.Controls.Add(btnNote, 2, 0); 
            tlp.Controls.Add(lbl, 3, 0);
            
            card.Controls.Add(tlp); 
            taskPanel.Controls.Add(card);
        }
    }

    public void AddNewTask(string name, string month, string date, string time, string note, string type) {
        tasks.Add(new RecurringTask() { 
            Name = name, MonthStr = month, DateStr = date, 
            TimeStr = time, Note = note, TaskType = type, LastTriggeredDate = "" 
        });
        SaveTasks(); 
        RefreshUI();
    }

    public void UpdateTask(int index, string name, string month, string date, string time, string note, string type) {
        if (index >= 0 && index < tasks.Count) {
            tasks[index].Name = name; tasks[index].MonthStr = month; tasks[index].DateStr = date; 
            tasks[index].TimeStr = time; tasks[index].Note = note; tasks[index].TaskType = type;
            SaveTasks(); 
            RefreshUI();
        }
    }

    public void DeleteTask(RecurringTask task) { 
        if (tasks.Contains(task)) { tasks.Remove(task); SaveTasks(); RefreshUI(); } 
    }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays, string sFreq) {
        digestType = dType; digestTimeStr = dTime; advanceDays = aDays; scanFrequency = sFreq;
        checkTimer.Enabled = false; checkTimer.Interval = GetTimerInterval(sFreq); checkTimer.Enabled = true;
        SaveTasks(); 
        MessageBox.Show("設定儲存成功！");
    }

    private void CheckTasks() {
        DateTime now = DateTime.Now; 
        bool needsSave = false;
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
                        t.LastTriggeredDate = targetDateStr; 
                        needsSave = true;
                        parentForm.AlertTab(1);
                        
                        if (t.TaskType == "單次" || t.TaskType == "到期日") {
                            toRemove.Add(t);
                        }
                    }
                }
            }
        }

        if (toRemove.Count > 0) {
            foreach (var r in toRemove) tasks.Remove(r);
            needsSave = true;
            this.Invoke(new Action(() => RefreshUI()));
        }

        if (digestType != "不提醒") {
            DateTime dtDigest;
            if (DateTime.TryParseExact(digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtDigest)) {
                DateTime targetDigest = new DateTime(now.Year, now.Month, now.Day, dtDigest.Hour, dtDigest.Minute, 0);
                bool shouldTrigger = false;
                if (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday && now >= targetDigest) shouldTrigger = true;
                if (digestType == "每月1號" && now.Day == 1 && now >= targetDigest) shouldTrigger = true;
                
                string todayStr = now.ToString("yyyy-MM-dd");
                if (shouldTrigger && lastDigestDate != todayStr) { 
                    lastDigestDate = todayStr; needsSave = true; new AllTasksViewWindow(this).Show(); 
                }
            }
        }

        if (needsSave) SaveTasks();
    }

    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target) {
        target = now;
        try {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]);
            int m = int.Parse(timeParts[1]);

            // 【新增邏輯】：如果是指定日期
            if (t.MonthStr == "特定日期") {
                if (DateTime.TryParseExact(t.DateStr, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime specificDate)) {
                    target = new DateTime(specificDate.Year, specificDate.Month, specificDate.Day, h, m, 0);
                    return true;
                }
                return false;
            }

            if (t.MonthStr == "每天") { 
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0); 
                if (now > target) target = target.AddDays(1); 
                return true; 
            }
            
            if (t.MonthStr == "每週") {
                Dictionary<string, DayOfWeek> dow = new Dictionary<string, DayOfWeek>();
                dow.Add("一", DayOfWeek.Monday); dow.Add("二", DayOfWeek.Tuesday); dow.Add("三", DayOfWeek.Wednesday);
                dow.Add("四", DayOfWeek.Thursday); dow.Add("五", DayOfWeek.Friday); dow.Add("六", DayOfWeek.Saturday); dow.Add("日", DayOfWeek.Sunday);
                if (!dow.ContainsKey(t.DateStr)) return false;
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                while (target.DayOfWeek != dow[t.DateStr] || now > target) target = target.AddDays(1);
                return true;
            }
            
            if (t.MonthStr == "每月" || t.MonthStr.EndsWith("月")) {
                int month = t.MonthStr == "每月" ? now.Month : int.Parse(t.MonthStr.Replace("月",""));
                int day = (t.DateStr == "月底") ? DateTime.DaysInMonth(now.Year, month) : int.Parse(t.DateStr);
                int validDay = Math.Min(day, DateTime.DaysInMonth(now.Year, month));
                
                target = new DateTime(now.Year, month, validDay, h, m, 0);
                if (now > target) {
                    if (t.MonthStr == "每月") target = target.AddMonths(1);
                    else target = target.AddYears(1);
                }
                return true;
            }
        } catch { } 
        return false;
    }

    public void SaveTasks() {
        List<string> lines = new List<string>();
        lines.Add(string.Format("#DIGEST|{0}|{1}|{2}|{3}|{4}", digestType, digestTimeStr, lastDigestDate, advanceDays, scanFrequency));
        foreach(var t in tasks) {
            lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate, EncodeBase64(t.Note), t.TaskType));
        }
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
                    LastTriggeredDate = p.Length > 4 ? p[4] : "", Note = p.Length > 5 ? DecodeBase64(p[5]) : "",
                    TaskType = p.Length > 6 ? p[6] : "循環"
                });
            }
        }
        RefreshUI();
    }

    private string ShowNoteEditBox(string name, string current) {
        Form f = new Form();
        f.Width = 400; f.Height = 350; f.Text = "編輯備註"; 
        f.StartPosition = FormStartPosition.CenterScreen; 
        f.FormBorderStyle = FormBorderStyle.FixedDialog;
        f.TopMost = true; // 保持上層

        Label lbl = new Label() { Text = "【" + name + "】", Left = 15, Top = 15, AutoSize = true, Font = new Font(MainFont, FontStyle.Bold) };
        TextBox txt = new TextBox() { Left = 15, Top = 50, Width = 350, Height = 180, Multiline = true, Text = current };
        Button btn = new Button() { Text = "儲存", Left = 265, Top = 250, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };

        f.Controls.AddRange(new Control[] { lbl, txt, btn });
        return f.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }

    private string EncodeBase64(string t) { return string.IsNullOrEmpty(t) ? "" : Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(t)); }
    private string DecodeBase64(string b) { try { return string.IsNullOrEmpty(b) ? "" : System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(b)); } catch { return ""; } }
}

// ==========================================
// 視窗：新增任務 (加入日期選擇器)
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD, cmType;
    private DateTimePicker dtpTime, dtpDate;
    private Label lblCycle, lblDate;

    public AddRecurringTaskWindow(App_RecurringTasks p) {
        this.parent = p; this.Text = "新增任務"; this.Width = 380; this.Height = 600; 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; // 【修正】保持獨立最上層，釋放主視窗操作權

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };

        f.Controls.Add(new Label() { Text = "任務名稱：" }); 
        txtN = new TextBox() { Width = 300 }; f.Controls.Add(txtN);

        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, 10, 0, 0) });
        txtNote = new TextBox() { Width = 300, Height = 60, Multiline = true }; f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, 10, 0, 0) });
        cmType = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmType.Items.Add("循環"); cmType.Items.Add("單次"); cmType.Items.Add("到期日");
        f.Controls.Add(cmType);

        lblCycle = new Label() { Text = "週期類型：", Margin = new Padding(0, 10, 0, 0) };
        f.Controls.Add(lblCycle);
        
        cmM = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.Add("每天"); cmM.Items.Add("每週"); cmM.Items.Add("每月");
        for(int i = 1; i <= 12; i++) cmM.Items.Add(i.ToString() + "月");
        f.Controls.Add(cmM); 
        
        cmD = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList }; 
        f.Controls.Add(cmD);

        // 【新增】指定日期控制項
        lblDate = new Label() { Text = "指定日期：", Margin = new Padding(0, 10, 0, 0) };
        f.Controls.Add(lblDate);
        dtpDate = new DateTimePicker() { Width = 300, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
        f.Controls.Add(dtpDate);

        // 【動態顯示邏輯】：循環顯示選單，單次顯示日曆
        cmType.SelectedIndexChanged += (s, e) => {
            bool isLoop = cmType.Text == "循環";
            lblCycle.Visible = cmM.Visible = cmD.Visible = isLoop;
            lblDate.Visible = dtpDate.Visible = !isLoop;
        };
        cmType.SelectedIndex = 0; 

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { cmD.Items.Add("每日"); cmD.Enabled = false; }
            else if(cmM.Text == "每週") { 
                cmD.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); cmD.Enabled = true; 
            }
            else { 
                for(int i = 1; i <= 31; i++) cmD.Items.Add(i.ToString()); 
                cmD.Items.Add("月底"); cmD.Enabled = true; 
            }
            if (cmD.Items.Count > 0) cmD.SelectedIndex = 0;
        }; 
        cmM.SelectedIndex = 0;

        f.Controls.Add(new Label() { Text = "觸發時間：", Margin = new Padding(0, 10, 0, 0) });
        dtpTime = new DateTimePicker() { Width = 300, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Value = DateTime.Today.AddHours(9) };
        f.Controls.Add(dtpTime);

        Button btn = new Button() { Text = "建立任務", Width = 300, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 20, 0, 0) };
        btn.Click += (s, e) => { 
            if(!string.IsNullOrWhiteSpace(txtN.Text)) { 
                string monthVal = cmType.Text == "循環" ? cmM.Text : "特定日期";
                string dateVal = cmType.Text == "循環" ? cmD.Text : dtpDate.Value.ToString("yyyy-MM-dd");
                parent.AddNewTask(txtN.Text, monthVal, dateVal, dtpTime.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); 
                this.Close(); 
            } 
        };
        f.Controls.Add(btn); 
        this.Controls.Add(f);
    }
}

// ==========================================
// 視窗：編輯任務
// ==========================================
public class EditRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private int idx;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD, cmType;
    private DateTimePicker dtpTime, dtpDate;
    private Label lblCycle, lblDate;

    public EditRecurringTaskWindow(App_RecurringTasks p, int i, App_RecurringTasks.RecurringTask t) {
        this.parent = p; this.idx = i; this.Text = "調整任務"; this.Width = 380; this.Height = 600; 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; // 【修正】

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };

        f.Controls.Add(new Label() { Text = "任務名稱：" }); 
        txtN = new TextBox() { Width = 300, Text = t.Name }; f.Controls.Add(txtN);

        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, 10, 0, 0) });
        txtNote = new TextBox() { Width = 300, Height = 60, Multiline = true, Text = t.Note }; f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, 10, 0, 0) });
        cmType = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmType.Items.Add("循環"); cmType.Items.Add("單次"); cmType.Items.Add("到期日");
        cmType.Text = t.TaskType; f.Controls.Add(cmType);

        lblCycle = new Label() { Text = "週期類型：", Margin = new Padding(0, 10, 0, 0) };
        f.Controls.Add(lblCycle);
        
        cmM = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.Add("每天"); cmM.Items.Add("每週"); cmM.Items.Add("每月");
        for(int k = 1; k <= 12; k++) cmM.Items.Add(k.ToString() + "月");
        f.Controls.Add(cmM);

        cmD = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList }; f.Controls.Add(cmD);

        lblDate = new Label() { Text = "指定日期：", Margin = new Padding(0, 10, 0, 0) };
        f.Controls.Add(lblDate);
        
        dtpDate = new DateTimePicker() { Width = 300, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
        f.Controls.Add(dtpDate);

        cmType.SelectedIndexChanged += (s, e) => {
            bool isLoop = cmType.Text == "循環";
            lblCycle.Visible = cmM.Visible = cmD.Visible = isLoop;
            lblDate.Visible = dtpDate.Visible = !isLoop;
        };

        if (t.MonthStr == "特定日期") {
            if (DateTime.TryParseExact(t.DateStr, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime d)) dtpDate.Value = d;
        } else {
            cmM.Text = t.MonthStr;
        }

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { cmD.Items.Add("每日"); }
            else if(cmM.Text == "每週") { cmD.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); }
            else { 
                for(int k = 1; k <= 31; k++) cmD.Items.Add(k.ToString()); 
                cmD.Items.Add("月底"); 
            }
            if(cmD.Items.Count > 0) cmD.SelectedIndex = 0;
        }; 
        if (t.MonthStr != "特定日期") cmD.Text = t.DateStr;

        f.Controls.Add(new Label() { Text = "觸發時間：", Margin = new Padding(0, 10, 0, 0) });
        dtpTime = new DateTimePicker() { Width = 300, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        if(DateTime.TryParseExact(t.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtpTime.Value = dtv;
        f.Controls.Add(dtpTime);

        Button btn = new Button() { Text = "儲存修改", Width = 300, Height = 40, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 20, 0, 0) };
        btn.Click += (s, e) => { 
            string monthVal = cmType.Text == "循環" ? cmM.Text : "特定日期";
            string dateVal = cmType.Text == "循環" ? cmD.Text : dtpDate.Value.ToString("yyyy-MM-dd");
            parent.UpdateTask(idx, txtN.Text, monthVal, dateVal, dtpTime.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); 
            this.Close(); 
        };
        f.Controls.Add(btn); 
        this.Controls.Add(f);
    }
}

// ==========================================
// 視窗：設定
// ==========================================
public class RecurringSettingsWindow : Form {
    private App_RecurringTasks parent;
    private ComboBox cmDig, cmAdv, cmScan;
    private DateTimePicker dtp;
    
    public RecurringSettingsWindow(App_RecurringTasks p) {
        this.parent = p; this.Text = "全域排程設定"; this.Width = 350; this.Height = 350; 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; // 【修正】

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };
        
        FlowLayoutPanel r1 = new FlowLayoutPanel() { AutoSize = true };
        r1.Controls.Add(new Label() { Text = "所有任務提前", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmAdv = new ComboBox() { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList }; 
        for (int i = 0; i <= 7; i++) cmAdv.Items.Add(i.ToString());
        cmAdv.Text = p.advanceDays.ToString(); 
        r1.Controls.Add(cmAdv); 
        r1.Controls.Add(new Label() { Text = "天加入待辦", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        f.Controls.Add(r1);
        
        FlowLayoutPanel r2 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 10, 0, 10) };
        r2.Controls.Add(new Label() { Text = "視窗摘要提醒：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmDig = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmDig.Items.Add("不提醒"); cmDig.Items.Add("每週一"); cmDig.Items.Add("每月1號");
        cmDig.Text = p.digestType;
        dtp = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        if(DateTime.TryParseExact(p.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtp.Value = dtv;
        r2.Controls.Add(cmDig); r2.Controls.Add(dtp); f.Controls.Add(r2);
        
        f.Controls.Add(new Label() { AutoSize = false, Height = 2, Width = 290, BorderStyle = BorderStyle.Fixed3D, Margin = new Padding(0, 5, 0, 15) });

        FlowLayoutPanel r3 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
        r3.Controls.Add(new Label() { Text = "待辦事項掃描頻率：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmScan = new ComboBox() { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
        cmScan.Items.AddRange(new string[] { "即時", "1分鐘", "5分鐘", "10分鐘", "1小時", "12小時", "1天" });
        cmScan.Text = p.scanFrequency;
        r3.Controls.Add(cmScan); f.Controls.Add(r3);

        Button btn = new Button() { Text = "儲存所有設定", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btn.Click += (s, e) => { 
            int.TryParse(cmAdv.Text, out int advDays);
            p.UpdateGlobalSettings(cmDig.Text, dtp.Value.ToString("HH:mm"), advDays, cmScan.Text); 
            this.Close(); 
        };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}

// ==========================================
// 視窗：全部檢視
// ==========================================
public class AllTasksViewWindow : Form {
    private App_RecurringTasks parentControl;
    private FlowLayoutPanel flow;

    public AllTasksViewWindow(App_RecurringTasks parent) {
        this.parentControl = parent; this.Text = "週期任務總覽"; this.Width = 820; this.Height = 800; 
        this.BackColor = Color.White;
        this.TopMost = true; // 【修正】

        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 60, BackColor = Color.WhiteSmoke, ColumnCount = 2 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170f));

        Label lbl = new Label() { Text = "週期任務排程總覽", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(15,0,0,0), Font = new Font("Microsoft JhengHei UI", 16f, FontStyle.Bold) };
        Button btnPrint = new Button() { Text = "轉存 PDF / 列印", Width = 150, Height = 35, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Right, Margin = new Padding(0,0,15,0) };
        header.Controls.Add(lbl, 0, 0); header.Controls.Add(btnPrint, 1, 0); this.Controls.Add(header);

        flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15, 20, 15, 15), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        flow.Resize += (s, e) => { 
            int w = flow.ClientSize.Width - 35; 
            if (w > 0) foreach (Control c in flow.Controls) if (c is GroupBox) c.Width = w;
        };
        this.Controls.Add(flow); flow.BringToFront(); RefreshData();
    }

    public void RefreshData() {
        flow.Controls.Clear(); 
        var tasks = parentControl.tasks;
        
        AddGroup("每天觸發", tasks.Where(t => t.MonthStr == "每天").ToList());
        AddGroup("每週觸發", tasks.Where(t => t.MonthStr == "每週").ToList());
        AddGroup("每月觸發", tasks.Where(t => t.MonthStr == "每月").ToList());
        for (int i = 1; i <= 12; i++) AddGroup(i.ToString() + "月 限定", tasks.Where(t => t.MonthStr == i.ToString() + "月").ToList());
        
        // 【新增】：顯示單次與到期日
        AddGroup("特定日期 (單次/到期日)", tasks.Where(t => t.MonthStr == "特定日期").ToList());
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
            bE.Click += (s, e) => { if(parentControl.tasks.IndexOf(t) != -1) new EditRecurringTaskWindow(parentControl, parentControl.tasks.IndexOf(t), t).Show(); };

            Button bD = new Button() { Text = "✕", Height = 28, Dock = DockStyle.Top, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            bD.Click += (s, e) => { if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { parentControl.DeleteTask(t); RefreshData(); } };

            Button bN = new Button() { Text = "註", Height = 28, Dock = DockStyle.Top, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold) };
            if (!string.IsNullOrEmpty(t.Note)) { bN.BackColor = Color.FromArgb(255,193,7); bN.ForeColor = Color.Black; } 
            else { bN.BackColor = Color.FromArgb(230,230,230); bN.ForeColor = Color.Gray; }
            bN.Click += (s, e) => { 
                Form nf = new Form() { Width = 400, Height = 350, Text = "任務備註", StartPosition = FormStartPosition.CenterScreen, TopMost = true };
                TextBox nt = new TextBox() { Left = 15, Top = 50, Width = 350, Height = 180, Multiline = true, Text = t.Note };
                Button nb = new Button() { Text = "儲存", Left = 265, Top = 250, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0,122,255), ForeColor = Color.White };
                nf.Controls.AddRange(new Control[] { new Label() { Text = "【" + t.Name + "】", Left = 15, Top = 15, AutoSize = true }, nt, nb });
                if(nf.ShowDialog() == DialogResult.OK) { t.Note = nt.Text; parentControl.SaveTasks(); RefreshData(); }
            };

            string typeTag = string.Format("[{0}] ", t.TaskType);
            string timeInfo = t.MonthStr == "特定日期" ? string.Format("[{0} {1}]", t.DateStr, t.TimeStr) : string.Format("[{0}] {1}", t.TimeStr, t.DateStr);

            row.Controls.Add(bE, 0, 0); row.Controls.Add(bD, 1, 0); row.Controls.Add(bN, 2, 0);
            row.Controls.Add(new Label() { Text = typeTag + timeInfo + "  " + t.Name, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Padding = new Padding(0,8,0,8) }, 3, 0);
            inner.Controls.Add(row);
        }
        gb.Controls.Add(inner); flow.Controls.Add(gb);
    }
}
