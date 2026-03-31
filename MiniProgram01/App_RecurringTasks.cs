using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_RecurringTasks : UserControl {
    private MainForm parentForm;
    private App_TodoList todoApp;
    private string recurringFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_recurring.txt");
    private FlowLayoutPanel taskPanel;
    private Timer checkTimer;

    // 全域設定變數
    public string digestType = "不提醒";
    public string digestTimeStr = "08:00";
    public string lastDigestDate = "";
    public int advanceDays = 0;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public class RecurringTask { public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate; }
    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; 
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        // 頂部控制列
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 3 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label() { Text = "已排程任務清單", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(10,0,0,0) };
        Button btnAdd = new Button() { Text = "新增任務", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand, BackColor = Color.White };
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this).ShowDialog(); };
        Button btnSet = new Button() { Text = "排程設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        header.Controls.AddRange(new Control[] { lblTitle, btnAdd, btnSet });
        this.Controls.Add(header);

        // 任務清單容器 (取代舊版的 ListBox)
        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadTasks();
        
        checkTimer = new Timer() { Interval = 600000, Enabled = true }; // 每 10 分鐘檢查一次
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks(); // 啟動時先檢查一次
    }

    // ==========================================
    // UI 更新與卡片生成 (加入調整與刪除按鈕)
    // ==========================================
    public void RefreshUI() {
        taskPanel.Controls.Clear();
        for (int i = 0; i < tasks.Count; i++) {
            var t = tasks[i];
            int currentIndex = i; 

            Panel card = new Panel() { Width = this.Width - 30, AutoSize = true, MinimumSize = new Size(0, 45), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(5, 5, 0, 5), BackColor = Color.FromArgb(248, 248, 250) };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, AutoSize = true, Padding = new Padding(5) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 55f)); // 最左側留給調整按鈕
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // 中間放文字
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); // 最右側放刪除按鈕

            // 【重點功能】前方的調整按鈕
            Button btnEdit = new Button() { Text = "調整", Width = 50, Height = 28, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0, 3, 0, 0) };
            btnEdit.Click += (s, e) => { new EditRecurringTaskWindow(this, currentIndex, t).ShowDialog(); };

            Label lbl = new Label() { Text = string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Font = MainFont };

            Button btnDel = new Button() { Text = "X", Width = 30, Height = 28, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(5, 3, 0, 0) };
            btnDel.Click += (s, e) => {
                if (MessageBox.Show("確定移除此任務？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) {
                    tasks.RemoveAt(currentIndex); SaveTasks(); RefreshUI();
                }
            };

            tlp.Controls.Add(btnEdit, 0, 0);
            tlp.Controls.Add(lbl, 1, 0);
            tlp.Controls.Add(btnDel, 2, 0);
            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    // ==========================================
    // API：提供子視窗呼叫
    // ==========================================
    public void AddNewTask(string name, string month, string date, string time) {
        tasks.Add(new RecurringTask() { Name = name, MonthStr = month, DateStr = date, TimeStr = time, LastTriggeredDate = "" });
        SaveTasks(); RefreshUI();
    }

    public void UpdateTask(int index, string name, string month, string date, string time) {
        if (index >= 0 && index < tasks.Count) {
            tasks[index].Name = name; tasks[index].MonthStr = month; tasks[index].DateStr = date; tasks[index].TimeStr = time;
            SaveTasks(); RefreshUI();
        }
    }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays) {
        digestType = dType; digestTimeStr = dTime; advanceDays = aDays;
        SaveTasks(); MessageBox.Show("設定已儲存！");
    }

    // ==========================================
    // 核心邏輯：檢查觸發條件
    // ==========================================
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
                        todoApp.AddTask(prefix + t.Name); // 推送至待辦
                        t.LastTriggeredDate = targetDateStr; // 鎖定記憶，防止重複觸發
                        needsSave = true;
                        parentForm.AlertTab(1); // 觸發待辦分頁閃爍
                    }
                }
            }
        }

        // 清單摘要提醒邏輯
        if (digestType != "不提醒") {
            DateTime dtDigest;
            if (DateTime.TryParseExact(digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtDigest)) {
                DateTime targetDigest = new DateTime(now.Year, now.Month, now.Day, dtDigest.Hour, dtDigest.Minute, 0);
                bool shouldTrigger = false;
                if (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday && now >= targetDigest) shouldTrigger = true;
                if (digestType == "每月1號" && now.Day == 1 && now >= targetDigest) shouldTrigger = true;

                string todayStr = now.ToString("yyyy-MM-dd");
                if (shouldTrigger && lastDigestDate != todayStr) {
                    todoApp.AddTask(string.Format("[系統提示] 目前共有 {0} 項週期任務正在背景為您守候中！", tasks.Count));
                    lastDigestDate = todayStr; needsSave = true; parentForm.AlertTab(1);
                }
            }
        }
        if (needsSave) SaveTasks();
    }

    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target) {
        target = now;
        try {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]); int m = int.Parse(timeParts[1]);

            if (t.MonthStr == "每天") {
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                if (now > target) target = target.AddDays(1);
                return true;
            } else if (t.MonthStr == "每週") {
                Dictionary<string, DayOfWeek> dow = new Dictionary<string, DayOfWeek>() {
                    {"一", DayOfWeek.Monday}, {"二", DayOfWeek.Tuesday}, {"三", DayOfWeek.Wednesday},
                    {"四", DayOfWeek.Thursday}, {"五", DayOfWeek.Friday}, {"六", DayOfWeek.Saturday}, {"日", DayOfWeek.Sunday}
                };
                if (!dow.ContainsKey(t.DateStr)) return false;
                DayOfWeek targetDow = dow[t.DateStr];
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                while (target.DayOfWeek != targetDow || (target.DayOfWeek == targetDow && now > target)) target = target.AddDays(1);
                return true;
            } else if (t.MonthStr == "每月") {
                int day = 1; bool isEndOfMonth = (t.DateStr == "月底");
                if (!isEndOfMonth && !int.TryParse(t.DateStr, out day)) return false;

                target = new DateTime(now.Year, now.Month, 1, h, m, 0);
                if (isEndOfMonth) target = target.AddMonths(1).AddDays(-1); 
                else target = target.AddDays(Math.Min(day, DateTime.DaysInMonth(now.Year, now.Month)) - 1);

                if (now > target) {
                    target = new DateTime(now.Year, now.Month, 1, h, m, 0).AddMonths(1);
                    if (isEndOfMonth) target = target.AddMonths(1).AddDays(-1);
                    else target = target.AddDays(Math.Min(day, DateTime.DaysInMonth(target.Year, target.Month)) - 1);
                }
                return true;
            }
        } catch { } return false;
    }

    private void SaveTasks() {
        List<string> lines = new List<string>(){ string.Format("#DIGEST|{0}|{1}|{2}|{3}", digestType, digestTimeStr, lastDigestDate, advanceDays) };
        foreach(var t in tasks) lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate));
        File.WriteAllLines(recurringFile, lines);
    }

    private void LoadTasks() {
        if(!File.Exists(recurringFile)) return;
        tasks.Clear();
        foreach(var l in File.ReadAllLines(recurringFile)) {
            if(l.StartsWith("#DIGEST")) { 
                var p = l.Split('|'); 
                digestType = p[1]; digestTimeStr = p[2]; lastDigestDate = p[3]; 
                if(p.Length >= 5) int.TryParse(p[4], out advanceDays); 
            } else {
                var p = l.Split('|');
                if(p.Length >= 4) tasks.Add(new RecurringTask() { Name = p[0], MonthStr = p[1], DateStr = p[2], TimeStr = p[3], LastTriggeredDate = p.Length > 4 ? p[4] : "" });
            }
        }
        RefreshUI();
    }
}

// ==========================================
// 視窗一：新增任務視窗
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parentControl;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate;
    private DateTimePicker dtpTime;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public AddRecurringTaskWindow(App_RecurringTasks parent) {
        this.parentControl = parent; this.Text = "新增週期任務";
        this.Width = 350; this.Height = 280; this.StartPosition = FormStartPosition.CenterParent; this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20) };
        
        flow.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        txtName = new TextBox() { Width = 280, Margin = new Padding(0, 0, 0, 15) }; flow.Controls.Add(txtName);

        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
        cmbMonth = new ComboBox() { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList }; cmbMonth.Items.AddRange(new string[] { "每天", "每週", "每月" });
        cmbDate = new ComboBox() { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList };
        dtpTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        
        cmbMonth.SelectedIndexChanged += (s, e) => {
            cmbDate.Items.Clear();
            if (cmbMonth.Text == "每天") { cmbDate.Items.Add("每日"); cmbDate.Enabled = false; } 
            else if (cmbMonth.Text == "每週") { cmbDate.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); cmbDate.Enabled = true; } 
            else if (cmbMonth.Text == "每月") { for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i.ToString()); cmbDate.Items.Add("月底"); cmbDate.Enabled = true; }
            cmbDate.SelectedIndex = 0;
        };
        cmbMonth.SelectedIndex = 0;
        
        row.Controls.AddRange(new Control[] { cmbMonth, cmbDate, dtpTime }); flow.Controls.Add(row);

        Button btnSave = new Button() { Text = "建立任務", Width = 280, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入任務名稱！"); return; }
            parentControl.AddNewTask(txtName.Text.Trim(), cmbMonth.Text, cmbDate.Text, dtpTime.Value.ToString("HH:mm"));
            this.Close();
        };
        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}

// ==========================================
// 視窗二：調整任務專屬視窗 (全新功能)
// ==========================================
public class EditRecurringTaskWindow : Form {
    private App_RecurringTasks parentControl;
    private int taskIndex;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate;
    private DateTimePicker dtpTime;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public EditRecurringTaskWindow(App_RecurringTasks parent, int index, App_RecurringTasks.RecurringTask task) {
        this.parentControl = parent; this.taskIndex = index; this.Text = "調整任務條件";
        this.Width = 350; this.Height = 280; this.StartPosition = FormStartPosition.CenterParent; this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20) };
        
        flow.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        txtName = new TextBox() { Width = 280, Text = task.Name, Margin = new Padding(0, 0, 0, 15) }; flow.Controls.Add(txtName);

        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
        cmbMonth = new ComboBox() { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList }; cmbMonth.Items.AddRange(new string[] { "每天", "每週", "每月" });
        cmbDate = new ComboBox() { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList };
        dtpTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        
        cmbMonth.SelectedIndexChanged += (s, e) => {
            cmbDate.Items.Clear();
            if (cmbMonth.Text == "每天") { cmbDate.Items.Add("每日"); cmbDate.Enabled = false; } 
            else if (cmbMonth.Text == "每週") { cmbDate.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); cmbDate.Enabled = true; } 
            else if (cmbMonth.Text == "每月") { for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i.ToString()); cmbDate.Items.Add("月底"); cmbDate.Enabled = true; }
            cmbDate.SelectedIndex = 0;
        };
        
        // 回填現有資料
        cmbMonth.Text = task.MonthStr;
        if (cmbDate.Items.Contains(task.DateStr)) cmbDate.Text = task.DateStr;
        DateTime dt; if(DateTime.TryParseExact(task.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dt)) dtpTime.Value = dt;

        row.Controls.AddRange(new Control[] { cmbMonth, cmbDate, dtpTime }); flow.Controls.Add(row);

        Button btnSave = new Button() { Text = "儲存修改", Width = 280, Height = 40, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入任務名稱！"); return; }
            parentControl.UpdateTask(taskIndex, txtName.Text.Trim(), cmbMonth.Text, cmbDate.Text, dtpTime.Value.ToString("HH:mm"));
            this.Close();
        };
        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}

// ==========================================
// 視窗三：全域排程設定 (清單提醒與提前天數)
// ==========================================
public class RecurringSettingsWindow : Form {
    private App_RecurringTasks parentControl;
    private ComboBox cmbDigest, cmbAdvanceDays;
    private DateTimePicker dtpDigestTime;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public RecurringSettingsWindow(App_RecurringTasks parent) {
        this.parentControl = parent; this.Text = "全域排程設定";
        this.Width = 350; this.Height = 230; this.StartPosition = FormStartPosition.CenterParent; this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20) };
        
        // 提前佈署設定
        FlowLayoutPanel advRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
        advRow.Controls.Add(new Label() { Text = "所有任務提前", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbAdvanceDays = new ComboBox() { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
        for (int i = 0; i <= 7; i++) cmbAdvanceDays.Items.Add(i.ToString());
        cmbAdvanceDays.Text = parent.advanceDays.ToString();
        advRow.Controls.Add(cmbAdvanceDays);
        advRow.Controls.Add(new Label() { Text = "天加入待辦", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        flow.Controls.Add(advRow);

        // 摘要提醒設定
        FlowLayoutPanel digestRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
        digestRow.Controls.Add(new Label() { Text = "摘要報告：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbDigest = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[]{"不提醒", "每週一", "每月1號"}); 
        cmbDigest.Text = parent.digestType;
        dtpDigestTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        DateTime dt; if(DateTime.TryParseExact(parent.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dt)) dtpDigestTime.Value = dt;
        digestRow.Controls.AddRange(new Control[]{ cmbDigest, dtpDigestTime });
        flow.Controls.Add(digestRow);

        Button btnSave = new Button() { Text = "儲存所有設定", Width = 280, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnSave.Click += (s, e) => {
            int adv = 0; int.TryParse(cmbAdvanceDays.Text, out adv);
            parentControl.UpdateGlobalSettings(cmbDigest.Text, dtpDigestTime.Value.ToString("HH:mm"), adv);
            this.Close();
        };
        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}
