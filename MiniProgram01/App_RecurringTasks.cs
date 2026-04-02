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
        btnViewAll.Click += (s, e) => OpenAllTasksView();

        Button btnAdd = new Button();
        btnAdd.Text = "新增任務";
        btnAdd.Dock = DockStyle.Fill;
        btnAdd.FlatStyle = FlatStyle.Flat;
        btnAdd.Margin = new Padding(2, 8, 2, 8);
        btnAdd.Cursor = Cursors.Hand;
        btnAdd.BackColor = Color.White;
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this).ShowDialog(); };

        Button btnSet = new Button();
        btnSet.Text = "設定";
        btnSet.Dock = DockStyle.Fill;
        btnSet.FlatStyle = FlatStyle.Flat;
        btnSet.BackColor = Color.Gainsboro;
        btnSet.Margin = new Padding(2, 8, 8, 8);
        btnSet.Cursor = Cursors.Hand;
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

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

    public void OpenAllTasksView() { 
        this.Invoke(new Action(() => { new AllTasksViewWindow(this).Show(); })); 
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
            btnEdit.Click += (s, e) => { 
                int idx = tasks.IndexOf(t); 
                if(idx != -1) { 
                    new EditRecurringTaskWindow(this, idx, t).ShowDialog(); 
                    RefreshUI(); 
                } 
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
            Label lbl = new Label();
            lbl.Text = typeTag + string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name);
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
            Name = name, 
            MonthStr = month, 
            DateStr = date, 
            TimeStr = time, 
            Note = note, 
            TaskType = type, 
            LastTriggeredDate = "" 
        });
        SaveTasks(); 
        RefreshUI();
    }

    public void UpdateTask(int index, string name, string month, string date, string time, string note, string type) {
        if (index >= 0 && index < tasks.Count) {
            tasks[index].Name = name; 
            tasks[index].MonthStr = month; 
            tasks[index].DateStr = date; 
            tasks[index].TimeStr = time; 
            tasks[index].Note = note; 
            tasks[index].TaskType = type;
            SaveTasks(); 
            RefreshUI();
        }
    }

    public void DeleteTask(RecurringTask task) { 
        if (tasks.Contains(task)) { 
            tasks.Remove(task); 
            SaveTasks(); 
            RefreshUI(); 
        } 
    }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays, string sFreq) {
        digestType = dType; 
        digestTimeStr = dTime; 
        advanceDays = aDays; 
        scanFrequency = sFreq;
        
        checkTimer.Enabled = false; 
        checkTimer.Interval = GetTimerInterval(sFreq); 
        checkTimer.Enabled = true;
        
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
            foreach (var r in toRemove) {
                tasks.Remove(r);
            }
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
                    lastDigestDate = todayStr; 
                    needsSave = true; 
                    OpenAllTasksView(); 
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

            if (t.MonthStr == "每天") { 
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0); 
                if (now > target) target = target.AddDays(1); 
                return true; 
            }
            
            if (t.MonthStr == "每週") {
                Dictionary<string, DayOfWeek> dow = new Dictionary<string, DayOfWeek>();
                dow.Add("一", DayOfWeek.Monday);
                dow.Add("二", DayOfWeek.Tuesday);
                dow.Add("三", DayOfWeek.Wednesday);
                dow.Add("四", DayOfWeek.Thursday);
                dow.Add("五", DayOfWeek.Friday);
                dow.Add("六", DayOfWeek.Saturday);
                dow.Add("日", DayOfWeek.Sunday);
                
                if (!dow.ContainsKey(t.DateStr)) return false;
                
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                while (target.DayOfWeek != dow[t.DateStr] || now > target) {
                    target = target.AddDays(1);
                }
                return true;
            }
            
            if (t.MonthStr == "每月" || t.MonthStr.EndsWith("月")) {
                int month = t.MonthStr == "每月" ? now.Month : int.Parse(t.MonthStr.Replace("月",""));
                int day = (t.DateStr == "月底") ? DateTime.DaysInMonth(now.Year, month) : int.Parse(t.DateStr);
                int validDay = Math.Min(day, DateTime.DaysInMonth(now.Year, month));
                
                target = new DateTime(now.Year, month, validDay, h, m, 0);
                if (now > target) {
                    if (t.MonthStr == "每月") {
                        target = target.AddMonths(1);
                    } else {
                        target = target.AddYears(1);
                    }
                }
                return true;
            }
        } catch { } 
        return false;
    }

    public void SaveTasks() {
        List<string> lines = new List<string>();
        string header = string.Format("#DIGEST|{0}|{1}|{2}|{3}|{4}", digestType, digestTimeStr, lastDigestDate, advanceDays, scanFrequency);
        lines.Add(header);
        
        foreach(var t in tasks) {
            string encodedNote = EncodeBase64(t.Note);
            lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate, encodedNote, t.TaskType));
        }
        File.WriteAllLines(recurringFile, lines);
    }

    private void LoadTasks() {
        if(!File.Exists(recurringFile)) return;
        tasks.Clear();
        string[] fileLines = File.ReadAllLines(recurringFile);
        
        foreach(var l in fileLines) {
            var p = l.Split('|');
            if(l.StartsWith("#DIGEST")) { 
                digestType = p[1]; 
                digestTimeStr = p[2]; 
                lastDigestDate = p[3]; 
                if(p.Length >= 5) advanceDays = int.Parse(p[4]);
                if(p.Length >= 6) scanFrequency = p[5];
            } else if(p.Length >= 4) {
                RecurringTask rt = new RecurringTask();
                rt.Name = p[0];
                rt.MonthStr = p[1];
                rt.DateStr = p[2];
                rt.TimeStr = p[3];
                rt.LastTriggeredDate = p.Length > 4 ? p[4] : "";
                rt.Note = p.Length > 5 ? DecodeBase64(p[5]) : "";
                rt.TaskType = p.Length > 6 ? p[6] : "循環";
                tasks.Add(rt);
            }
        }
        RefreshUI();
    }

    private string ShowNoteEditBox(string name, string current) {
        Form f = new Form();
        f.Width = 400; 
        f.Height = 350; 
        f.Text = "編輯備註"; 
        f.StartPosition = FormStartPosition.CenterScreen; 
        f.FormBorderStyle = FormBorderStyle.FixedDialog;

        Label lbl = new Label();
        lbl.Text = "【" + name + "】";
        lbl.Left = 15; 
        lbl.Top = 15; 
        lbl.AutoSize = true; 
        lbl.Font = new Font(MainFont, FontStyle.Bold);

        TextBox txt = new TextBox();
        txt.Left = 15; 
        txt.Top = 50; 
        txt.Width = 350; 
        txt.Height = 180; 
        txt.Multiline = true; 
        txt.Text = current;

        Button btn = new Button();
        btn.Text = "儲存"; 
        btn.Left = 265; 
        btn.Top = 250; 
        btn.Width = 100; 
        btn.Height = 35; 
        btn.DialogResult = DialogResult.OK; 
        btn.FlatStyle = FlatStyle.Flat; 
        btn.BackColor = AppleBlue; 
        btn.ForeColor = Color.White;

        f.Controls.Add(lbl);
        f.Controls.Add(txt);
        f.Controls.Add(btn);

        if (f.ShowDialog() == DialogResult.OK) {
            return txt.Text;
        }
        return null;
    }

    private string EncodeBase64(string t) {
        if (string.IsNullOrEmpty(t)) return "";
        return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(t));
    }
    
    private string DecodeBase64(string b) {
        if (string.IsNullOrEmpty(b)) return "";
        try { 
            return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(b)); 
        } catch { 
            return ""; 
        } 
    }
}

// ==========================================
// 視窗：新增任務
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private TextBox txtN;
    private TextBox txtNote;
    private ComboBox cmM;
    private ComboBox cmD;
    private ComboBox cmType;
    private DateTimePicker dtp;

    public AddRecurringTaskWindow(App_RecurringTasks p) {
        this.parent = p; 
        this.Text = "新增任務"; 
        this.Width = 380; 
        this.Height = 550; 
        this.StartPosition = FormStartPosition.CenterScreen;

        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill; 
        f.FlowDirection = FlowDirection.TopDown; 
        f.Padding = new Padding(25);

        f.Controls.Add(new Label() { Text = "任務名稱：" }); 
        txtN = new TextBox() { Width = 300 }; 
        f.Controls.Add(txtN);

        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, 10, 0, 0) });
        txtNote = new TextBox() { Width = 300, Height = 60, Multiline = true }; 
        f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, 10, 0, 0) });
        cmType = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmType.Items.Add("循環");
        cmType.Items.Add("單次");
        cmType.Items.Add("到期日");
        cmType.SelectedIndex = 0; 
        f.Controls.Add(cmType);

        f.Controls.Add(new Label() { Text = "週期類型：", Margin = new Padding(0, 10, 0, 0) });
        cmM = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.Add("每天");
        cmM.Items.Add("每週");
        cmM.Items.Add("每月");
        for(int i = 1; i <= 12; i++) {
            cmM.Items.Add(i.ToString() + "月");
        }
        f.Controls.Add(cmM); 
        
        cmD = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList }; 
        f.Controls.Add(cmD);

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { 
                cmD.Items.Add("每日"); 
                cmD.Enabled = false; 
            }
            else if(cmM.Text == "每週") { 
                cmD.Items.Add("一");
                cmD.Items.Add("二");
                cmD.Items.Add("三");
                cmD.Items.Add("四");
                cmD.Items.Add("五");
                cmD.Items.Add("六");
                cmD.Items.Add("日");
                cmD.Enabled = true; 
            }
            else { 
                for(int i = 1; i <= 31; i++) {
                    cmD.Items.Add(i.ToString()); 
                }
                cmD.Items.Add("月底"); 
                cmD.Enabled = true; 
            }
            if (cmD.Items.Count > 0) cmD.SelectedIndex = 0;
        }; 
        cmM.SelectedIndex = 0;

        dtp = new DateTimePicker();
        dtp.Width = 300; 
        dtp.Format = DateTimePickerFormat.Custom; 
        dtp.CustomFormat = "HH:mm"; 
        dtp.ShowUpDown = true; 
        dtp.Value = DateTime.Today.AddHours(9); 
        dtp.Margin = new Padding(0, 10, 0, 0);
        f.Controls.Add(dtp);

        Button btn = new Button();
        btn.Text = "建立任務"; 
        btn.Width = 300; 
        btn.Height = 40; 
        btn.BackColor = Color.FromArgb(0, 122, 255); 
        btn.ForeColor = Color.White; 
        btn.FlatStyle = FlatStyle.Flat; 
        btn.Margin = new Padding(0, 20, 0, 0);
        btn.Click += (s, e) => { 
            if(!string.IsNullOrWhiteSpace(txtN.Text)) { 
                parent.AddNewTask(txtN.Text, cmM.Text, cmD.Text, dtp.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); 
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
    private TextBox txtN;
    private TextBox txtNote;
    private ComboBox cmM;
    private ComboBox cmD;
    private ComboBox cmType;
    private DateTimePicker dtp;

    public EditRecurringTaskWindow(App_RecurringTasks p, int i, App_RecurringTasks.RecurringTask t) {
        this.parent = p; 
        this.idx = i; 
        this.Text = "調整任務"; 
        this.Width = 380; 
        this.Height = 550; 
        this.StartPosition = FormStartPosition.CenterScreen;

        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill; 
        f.FlowDirection = FlowDirection.TopDown; 
        f.Padding = new Padding(25);

        f.Controls.Add(new Label() { Text = "任務名稱：" }); 
        txtN = new TextBox() { Width = 300, Text = t.Name }; 
        f.Controls.Add(txtN);

        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, 10, 0, 0) });
        txtNote = new TextBox() { Width = 300, Height = 60, Multiline = true, Text = t.Note }; 
        f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, 10, 0, 0) });
        cmType = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmType.Items.Add("循環");
        cmType.Items.Add("單次");
        cmType.Items.Add("到期日");
        cmType.Text = t.TaskType; 
        f.Controls.Add(cmType);

        f.Controls.Add(new Label() { Text = "週期類型：", Margin = new Padding(0, 10, 0, 0) });
        cmM = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
        cmM.Items.Add("每天");
        cmM.Items.Add("每週");
        cmM.Items.Add("每月");
        for(int k = 1; k <= 12; k++) {
            cmM.Items.Add(k.ToString() + "月");
        }
        cmM.Text = t.MonthStr; 
        f.Controls.Add(cmM);

        cmD = new ComboBox() { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList }; 
        f.Controls.Add(cmD);

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { 
                cmD.Items.Add("每日"); 
            }
            else if(cmM.Text == "每週") { 
                cmD.Items.Add("一");
                cmD.Items.Add("二");
                cmD.Items.Add("三");
                cmD.Items.Add("四");
                cmD.Items.Add("五");
                cmD.Items.Add("六");
                cmD.Items.Add("日");
            }
            else { 
                for(int k = 1; k <= 31; k++) {
                    cmD.Items.Add(k.ToString()); 
                }
                cmD.Items.Add("月底"); 
            }
            if(cmD.Items.Count > 0) cmD.SelectedIndex = 0;
        }; 
        cmD.Text = t.DateStr;

        dtp = new DateTimePicker();
        dtp.Width = 300; 
        dtp.Format = DateTimePickerFormat.Custom; 
        dtp.CustomFormat = "HH:mm"; 
        dtp.ShowUpDown = true;
        dtp.Margin = new Padding(0, 10, 0, 0);
        
        DateTime dtv;
        if(DateTime.TryParseExact(t.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtv)) {
            dtp.Value = dtv;
        }
        f.Controls.Add(dtp);

        Button btn = new Button();
        btn.Text = "儲存修改"; 
        btn.Width = 300; 
        btn.Height = 40; 
        btn.BackColor = Color.FromArgb(0, 153, 76); 
        btn.ForeColor = Color.White; 
        btn.FlatStyle = FlatStyle.Flat; 
        btn.Margin = new Padding(0, 20, 0, 0);
        
        btn.Click += (s, e) => { 
            parent.UpdateTask(idx, txtN.Text, cmM.Text, cmD.Text, dtp.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); 
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
    private ComboBox cmDig;
    private ComboBox cmAdv;
    private ComboBox cmScan;
    private DateTimePicker dtp;
    
    public RecurringSettingsWindow(App_RecurringTasks p) {
        this.parent = p; 
        this.Text = "全域排程設定"; 
        this.Width = 350; 
        this.Height = 350; 
        this.StartPosition = FormStartPosition.CenterScreen;
        
        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill; 
        f.FlowDirection = FlowDirection.TopDown; 
        f.Padding = new Padding(20);
        
        FlowLayoutPanel r1 = new FlowLayoutPanel();
        r1.AutoSize = true;
        r1.Controls.Add(new Label() { Text = "所有任務提前", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        
        cmAdv = new ComboBox();
        cmAdv.Width = 60; 
        cmAdv.DropDownStyle = ComboBoxStyle.DropDownList; 
        for (int i = 0; i <= 7; i++) {
            cmAdv.Items.Add(i.ToString());
        }
        cmAdv.Text = p.advanceDays.ToString(); 
        
        r1.Controls.Add(cmAdv); 
        r1.Controls.Add(new Label() { Text = "天加入待辦", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        f.Controls.Add(r1);
        
        FlowLayoutPanel r2 = new FlowLayoutPanel();
        r2.AutoSize = true; 
        r2.Margin = new Padding(0, 10, 0, 10);
        r2.Controls.Add(new Label() { Text = "視窗摘要提醒：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        
        cmDig = new ComboBox();
        cmDig.Width = 80; 
        cmDig.DropDownStyle = ComboBoxStyle.DropDownList; 
        cmDig.Items.Add("不提醒");
        cmDig.Items.Add("每週一");
        cmDig.Items.Add("每月1號");
        cmDig.Text = p.digestType;
        
        dtp = new DateTimePicker();
        dtp.Width = 80; 
        dtp.Format = DateTimePickerFormat.Custom; 
        dtp.CustomFormat = "HH:mm"; 
        dtp.ShowUpDown = true;
        
        DateTime dtv; 
        if(DateTime.TryParseExact(p.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtv)) {
            dtp.Value = dtv;
        }
        
        r2.Controls.Add(cmDig);
        r2.Controls.Add(dtp); 
        f.Controls.Add(r2);
        
        Label sep = new Label();
        sep.AutoSize = false; 
        sep.Height = 2; 
        sep.Width = 290; 
        sep.BorderStyle = BorderStyle.Fixed3D; 
        sep.Margin = new Padding(0, 5, 0, 15);
        f.Controls.Add(sep);

        FlowLayoutPanel r3 = new FlowLayoutPanel();
        r3.AutoSize = true; 
        r3.Margin = new Padding(0, 0, 0, 20);
        r3.Controls.Add(new Label() { Text = "待辦事項掃描頻率：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        
        cmScan = new ComboBox();
        cmScan.Width = 100; 
        cmScan.DropDownStyle = ComboBoxStyle.DropDownList;
        cmScan.Items.Add("即時");
        cmScan.Items.Add("1分鐘");
        cmScan.Items.Add("5分鐘");
        cmScan.Items.Add("10分鐘");
        cmScan.Items.Add("1小時");
        cmScan.Items.Add("12小時");
        cmScan.Items.Add("1天");
        cmScan.Text = p.scanFrequency;
        
        r3.Controls.Add(cmScan);
        f.Controls.Add(r3);

        Button btn = new Button();
        btn.Text = "儲存所有設定"; 
        btn.Width = 290; 
        btn.Height = 40; 
        btn.BackColor = Color.FromArgb(0, 122, 255); 
        btn.ForeColor = Color.White; 
        btn.FlatStyle = FlatStyle.Flat;
        btn.Click += (s, e) => { 
            int advDays = 0;
            int.TryParse(cmAdv.Text, out advDays);
            p.UpdateGlobalSettings(cmDig.Text, dtp.Value.ToString("HH:mm"), advDays, cmScan.Text); 
            this.Close(); 
        };
        f.Controls.Add(btn); 
        
        this.Controls.Add(f);
    }
}

// ==========================================
// 視窗：全部檢視
// ==========================================
public class AllTasksViewWindow : Form {
    private App_RecurringTasks parentControl;
    private FlowLayoutPanel flow;

    public AllTasksViewWindow(App_RecurringTasks parent) {
        this.parentControl = parent; 
        this.Text = "週期任務總覽"; 
        this.Width = 820; 
        this.Height = 800; 
        this.BackColor = Color.White;

        TableLayoutPanel header = new TableLayoutPanel();
        header.Dock = DockStyle.Top; 
        header.Height = 60; 
        header.BackColor = Color.WhiteSmoke; 
        header.ColumnCount = 2;
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170f));

        Label lbl = new Label();
        lbl.Text = "週期任務排程總覽"; 
        lbl.Dock = DockStyle.Fill; 
        lbl.TextAlign = ContentAlignment.MiddleLeft; 
        lbl.Padding = new Padding(15,0,0,0); 
        lbl.Font = new Font("Microsoft JhengHei UI", 16f, FontStyle.Bold);

        Button btnPrint = new Button();
        btnPrint.Text = "轉存 PDF / 列印"; 
        btnPrint.Width = 150; 
        btnPrint.Height = 35; 
        btnPrint.BackColor = Color.FromArgb(0, 153, 76); 
        btnPrint.ForeColor = Color.White; 
        btnPrint.FlatStyle = FlatStyle.Flat; 
        btnPrint.Anchor = AnchorStyles.Right; 
        btnPrint.Margin = new Padding(0,0,15,0);
        
        header.Controls.Add(lbl, 0, 0); 
        header.Controls.Add(btnPrint, 1, 0); 
        this.Controls.Add(header);

        flow = new FlowLayoutPanel();
        flow.Dock = DockStyle.Fill; 
        flow.AutoScroll = true; 
        flow.Padding = new Padding(15, 20, 15, 15); 
        flow.FlowDirection = FlowDirection.TopDown; 
        flow.WrapContents = false;

        flow.Resize += (s, e) => { 
            int w = flow.ClientSize.Width - 35; 
            if (w > 0) {
                foreach (Control c in flow.Controls) {
                    if (c is GroupBox) {
                        c.Width = w;
                    }
                }
            }
        };
        
        this.Controls.Add(flow); 
        flow.BringToFront(); 
        RefreshData();
    }

    public void RefreshData() {
        flow.Controls.Clear(); 
        var tasks = parentControl.tasks;
        
        AddGroup("每天觸發", tasks.Where(t => t.MonthStr == "每天").ToList());
        AddGroup("每週觸發", tasks.Where(t => t.MonthStr == "每週").ToList());
        AddGroup("每月觸發", tasks.Where(t => t.MonthStr == "每月").ToList());
        for (int i = 1; i <= 12; i++) {
            AddGroup(i.ToString() + "月 限定", tasks.Where(t => t.MonthStr == i.ToString() + "月").ToList());
        }
    }

    private void AddGroup(string header, List<App_RecurringTasks.RecurringTask> sub) {
        if (sub.Count == 0) return;

        GroupBox gb = new GroupBox();
        gb.Text = "【 " + header + " 】"; 
        gb.Font = new Font("Microsoft JhengHei UI", 12f, FontStyle.Bold); 
        gb.ForeColor = Color.FromArgb(0, 122, 255); 
        gb.AutoSize = true; 
        gb.Width = flow.ClientSize.Width - 35; 
        gb.Margin = new Padding(10, 10, 10, 25); 
        gb.Padding = new Padding(15, 25, 15, 15);

        FlowLayoutPanel inner = new FlowLayoutPanel();
        inner.Dock = DockStyle.Fill; 
        inner.FlowDirection = FlowDirection.TopDown; 
        inner.WrapContents = false; 
        inner.AutoSize = true;

        foreach (var t in sub) {
            TableLayoutPanel row = new TableLayoutPanel();
            row.Width = gb.Width - 40; 
            row.AutoSize = true; 
            row.ColumnCount = 4; 
            row.Margin = new Padding(0,0,0,8);
            
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button bE = new Button();
            bE.Text = "調"; 
            bE.Height = 28; 
            bE.Dock = DockStyle.Top; 
            bE.BackColor = Color.FromArgb(0, 122, 255); 
            bE.ForeColor = Color.White; 
            bE.FlatStyle = FlatStyle.Flat;
            bE.Click += (s, e) => { 
                if(parentControl.tasks.IndexOf(t) != -1) { 
                    new EditRecurringTaskWindow(parentControl, parentControl.tasks.IndexOf(t), t).ShowDialog(); 
                    RefreshData(); 
                } 
            };

            Button bD = new Button();
            bD.Text = "✕"; 
            bD.Height = 28; 
            bD.Dock = DockStyle.Top; 
            bD.BackColor = Color.IndianRed; 
            bD.ForeColor = Color.White; 
            bD.FlatStyle = FlatStyle.Flat;
            bD.Click += (s, e) => { 
                if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { 
                    parentControl.DeleteTask(t); 
                    RefreshData(); 
                } 
            };

            Button bN = new Button();
            bN.Text = "註"; 
            bN.Height = 28; 
            bN.Dock = DockStyle.Top; 
            bN.FlatStyle = FlatStyle.Flat; 
            bN.Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold);
            
            if (!string.IsNullOrEmpty(t.Note)) { 
                bN.BackColor = Color.FromArgb(255,193,7); 
                bN.ForeColor = Color.Black; 
            } else { 
                bN.BackColor = Color.FromArgb(230,230,230); 
                bN.ForeColor = Color.Gray; 
            }
            
            bN.Click += (s, e) => { 
                Form nf = new Form();
                nf.Width = 400; 
                nf.Height = 350; 
                nf.Text = "任務備註"; 
                nf.StartPosition = FormStartPosition.CenterScreen;
                
                Label lblName = new Label();
                lblName.Text = "【" + t.Name + "】"; 
                lblName.Left = 15; 
                lblName.Top = 15; 
                lblName.AutoSize = true;

                TextBox nt = new TextBox();
                nt.Left = 15; 
                nt.Top = 50; 
                nt.Width = 350; 
                nt.Height = 180; 
                nt.Multiline = true; 
                nt.Text = t.Note;

                Button nb = new Button();
                nb.Text = "儲存"; 
                nb.Left = 265; 
                nb.Top = 250; 
                nb.Width = 100; 
                nb.Height = 35; 
                nb.DialogResult = DialogResult.OK; 
                nb.FlatStyle = FlatStyle.Flat; 
                nb.BackColor = Color.FromArgb(0,122,255); 
                nb.ForeColor = Color.White;

                nf.Controls.Add(lblName);
                nf.Controls.Add(nt);
                nf.Controls.Add(nb);

                if(nf.ShowDialog() == DialogResult.OK) { 
                    t.Note = nt.Text; 
                    parentControl.SaveTasks(); 
                    RefreshData(); 
                }
            };

            Label lblTaskInfo = new Label();
            string typeTag = string.Format("[{0}] ", t.TaskType);
            lblTaskInfo.Text = typeTag + string.Format("[{0}] {1}  {2}", t.TimeStr, t.DateStr, t.Name);
            lblTaskInfo.Dock = DockStyle.Fill; 
            lblTaskInfo.TextAlign = ContentAlignment.MiddleLeft; 
            lblTaskInfo.AutoSize = true; 
            lblTaskInfo.Padding = new Padding(0,8,0,8);

            row.Controls.Add(bE, 0, 0); 
            row.Controls.Add(bD, 1, 0); 
            row.Controls.Add(bN, 2, 0);
            row.Controls.Add(lblTaskInfo, 3, 0);
            
            inner.Controls.Add(row);
        }
        gb.Controls.Add(inner); 
        flow.Controls.Add(gb);
    }
}
