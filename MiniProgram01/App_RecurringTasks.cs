// ============================================================
// FILE: MiniProgram01/App_RecurringTasks.cs 
// ============================================================
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Drawing.Printing; // 列印功能

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

    // --- iOS 風格色彩與字體 ---
    private static Color iosBackground = Color.FromArgb(242, 242, 247);
    private static Color iosCardWhite = Color.White;
    private static Color iosAppleBlue = Color.FromArgb(0, 122, 255);
    private static Color iosRed = Color.FromArgb(255, 59, 48);
    private static Color iosGray = Color.FromArgb(142, 142, 147);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    public class RecurringTask { 
        public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate, Note, TaskType; 
    }
    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; 
        this.todoApp = todoApp;
        this.BackColor = iosBackground;
        this.Padding = new Padding(15);
        this.AutoScaleMode = AutoScaleMode.Dpi; // 核心：支援高 DPI 縮放

        InitializeUI();
        LoadTasks();

        // 啟動背景檢查計時器
        checkTimer = new Timer() { Interval = GetTimerInterval(scanFrequency), Enabled = true };
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks();
    }

    private void InitializeUI() {
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 50, ColumnCount = 4 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));

        Label lblTitle = new Label() { Text = "週期任務", Font = new Font("Microsoft JhengHei UI", 14f, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5, 0, 0, 0) };
        Button btnViewAll = CreateIOSButton("全部", Color.FromArgb(230, 230, 235), iosGray);
        btnViewAll.Click += (s, e) => { new AllTasksViewWindow(this).Show(); }; 
        Button btnAdd = CreateIOSButton("新增", iosAppleBlue, Color.White);
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this, -1, null).ShowDialog(); };
        Button btnSet = CreateIOSButton("設定", Color.FromArgb(230, 230, 235), iosGray);
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnViewAll, 1, 0);
        header.Controls.Add(btnAdd, 2, 0);
        header.Controls.Add(btnSet, 3, 0);
        this.Controls.Add(header);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = iosBackground, Padding = new Padding(0, 10, 0, 0) };
        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - 25;
            if (safeWidth > 0) foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
        };
        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();
    }

    private Button CreateIOSButton(string text, Color backColor, Color foreColor) {
        Button btn = new Button() { Text = text, Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2, 8, 2, 8), Cursor = Cursors.Hand, BackColor = backColor, ForeColor = foreColor, Font = BoldFont };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
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
            Panel card = new Panel() { Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 55), Margin = new Padding(5, 5, 5, 10), BackColor = iosCardWhite, BorderStyle = BorderStyle.None };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, AutoSize = true, Padding = new Padding(10) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 

            Button btnDel = CreateIconButton("✕", iosRed, Color.White);
            btnDel.Click += (s, e) => { 
                if (MessageBox.Show("確定移除週期任務？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK) {
                    tasks.Remove(t); SaveTasks(); RefreshUI();
                }
            };

            Button btnEdit = CreateIconButton("修", iosAppleBlue, Color.White);
            btnEdit.Click += (s, e) => { new AddRecurringTaskWindow(this, tasks.IndexOf(t), t).ShowDialog(); };

            Button btnRun = CreateIconButton("執", Color.FromArgb(52, 199, 89), Color.White);
            btnRun.Click += (s, e) => {
                todoApp.AddTask(t.Name, "Black", t.TaskType, t.Note);
                t.LastTriggeredDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                SaveTasks(); RefreshUI();
                MessageBox.Show("已手動觸發任務寫入待辦！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            string ruleDesc = $"每月 {t.MonthStr} 月 {t.DateStr} 號 {t.TimeStr} 觸發";
            Label lbl = new Label() { Text = $"{t.Name}\n({ruleDesc})", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Font = MainFont, Padding = new Padding(10, 0, 0, 0) };

            tlp.Controls.Add(btnDel, 0, 0);
            tlp.Controls.Add(btnEdit, 1, 0);
            tlp.Controls.Add(lbl, 2, 0);
            tlp.Controls.Add(btnRun, 3, 0);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    private Button CreateIconButton(string text, Color backColor, Color foreColor) {
        Button btn = new Button() { Text = text, Dock = DockStyle.Fill, Height = 35, BackColor = backColor, ForeColor = foreColor, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(2), Font = BoldFont };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    // --- 核心：檢查週期條件寫入待辦 ---
    public void CheckTasks() {
        DateTime now = DateTime.Now;
        bool triggeredAny = false;

        foreach (var t in tasks) {
            if (t.LastTriggeredDate.StartsWith(now.ToString("yyyy-MM-dd"))) continue; // 今日已觸發過
            
            bool matchMonth = t.MonthStr == "*" || t.MonthStr.Split(',').Contains(now.Month.ToString());
            bool matchDate = t.DateStr == "*" || t.DateStr.Split(',').Contains(now.Day.ToString());
            
            string[] timeParts = t.TimeStr.Split(':');
            bool matchTime = true;
            if (timeParts.Length == 2) {
                int targetH = int.Parse(timeParts[0]);
                int targetM = int.Parse(timeParts[1]);
                if (now.Hour < targetH || (now.Hour == targetH && now.Minute < targetM)) matchTime = false;
            }

            if (matchMonth && matchDate && matchTime) {
                todoApp.AddTask(t.Name, "Black", t.TaskType, t.Note);
                t.LastTriggeredDate = now.ToString("yyyy-MM-dd HH:mm");
                triggeredAny = true;
            }
        }

        if (triggeredAny) {
            SaveTasks();
            RefreshUI();
            if (parentForm != null) parentForm.Invoke(new Action(() => parentForm.AlertTab(1))); // 提醒待辦分頁
        }
    }

    // --- 存檔與載入 ---
    public void SaveTasks() {
        List<string> lines = new List<string>();
        lines.Add($"CONFIG|{digestType}|{digestTimeStr}|{lastDigestDate}|{advanceDays}|{scanFrequency}");
        foreach (var t in tasks) {
            string safeNote = string.IsNullOrEmpty(t.Note) ? "" : Convert.ToBase64String(Encoding.UTF8.GetBytes(t.Note));
            lines.Add($"{t.Name}|{t.MonthStr}|{t.DateStr}|{t.TimeStr}|{t.LastTriggeredDate}|{t.TaskType}|{safeNote}");
        }
        File.WriteAllLines(recurringFile, lines);
    }

    public void LoadTasks() {
        if (!File.Exists(recurringFile)) return;
        tasks.Clear();
        foreach (var l in File.ReadAllLines(recurringFile)) {
            var p = l.Split('|');
            if (p[0] == "CONFIG" && p.Length >= 6) {
                digestType = p[1]; digestTimeStr = p[2]; lastDigestDate = p[3];
                int.TryParse(p[4], out int adv); advanceDays = adv;
                scanFrequency = p[5];
                continue;
            }
            if (p.Length >= 7) {
                string note = "";
                try { note = string.IsNullOrEmpty(p[6]) ? "" : Encoding.UTF8.GetString(Convert.FromBase64String(p[6])); } catch { note = p[6]; }
                tasks.Add(new RecurringTask() { Name = p[0], MonthStr = p[1], DateStr = p[2], TimeStr = p[3], LastTriggeredDate = p[4], TaskType = p[5], Note = note });
            }
        }
        RefreshUI();
    }
}

// ==========================================
// 視窗：新增/編輯週期任務
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private int index;
    private TextBox txtName, txtMonth, txtDate, txtTime, txtType, txtNote;

    public AddRecurringTaskWindow(App_RecurringTasks p, int idx, App_RecurringTasks.RecurringTask item) {
        this.parent = p; this.index = idx;
        this.Text = idx == -1 ? "新增週期任務" : "編輯週期任務";
        this.Width = 450; this.Height = 520;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.BackColor = Color.White;
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10.5f);

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };
        
        f.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true, ForeColor = Color.FromArgb(142, 142, 147) });
        txtName = new TextBox() { Width = 380, Text = item?.Name ?? "", Margin = new Padding(0, 0, 0, 15), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtName);

        f.Controls.Add(new Label() { Text = "觸發月份 (* 代表每月, 逗號分隔如 1,3,5)：", AutoSize = true, ForeColor = Color.FromArgb(142, 142, 147) });
        txtMonth = new TextBox() { Width = 380, Text = item?.MonthStr ?? "*", Margin = new Padding(0, 0, 0, 15), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtMonth);

        f.Controls.Add(new Label() { Text = "觸發日期 (* 代表每日, 逗號分隔如 1,15)：", AutoSize = true, ForeColor = Color.FromArgb(142, 142, 147) });
        txtDate = new TextBox() { Width = 380, Text = item?.DateStr ?? "*", Margin = new Padding(0, 0, 0, 15), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtDate);

        f.Controls.Add(new Label() { Text = "觸發時間 (格式 HH:mm, 例如 08:30)：", AutoSize = true, ForeColor = Color.FromArgb(142, 142, 147) });
        txtTime = new TextBox() { Width = 380, Text = item?.TimeStr ?? "08:00", Margin = new Padding(0, 0, 0, 15), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtTime);

        Button btnSave = new Button() { Text = "儲存設定", Width = 380, Height = 42, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtName.Text)) return;
            var newTask = new App_RecurringTasks.RecurringTask() { 
                Name = txtName.Text, MonthStr = txtMonth.Text, DateStr = txtDate.Text, TimeStr = txtTime.Text, LastTriggeredDate = item?.LastTriggeredDate ?? "", TaskType = "週期", Note = "" 
            };
            if (index == -1) parent.tasks.Add(newTask);
            else parent.tasks[index] = newTask;
            parent.SaveTasks(); parent.RefreshUI(); this.Close();
        };
        f.Controls.Add(btnSave);
        this.Controls.Add(f);
    }
}

// ==========================================
// 視窗：系統設定與全部檢視 (省略細部排版，保留主邏輯)
// ==========================================
public class RecurringSettingsWindow : Form {
    public RecurringSettingsWindow(App_RecurringTasks p) {
        this.Text = "週期任務設定";
        this.Width = 350; this.Height = 200;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White;
        // 介面與參數寫入 p.scanFrequency 等... 略
    }
}

public class AllTasksViewWindow : Form {
    public AllTasksViewWindow(App_RecurringTasks p) {
        this.Text = "全部任務檢視與列印";
        this.Width = 600; this.Height = 400;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White;
        // 列印功能調用 System.Drawing.Printing 等... 略
    }
}
