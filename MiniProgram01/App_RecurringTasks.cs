using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_RecurringTasks : UserControl {
    private MainForm parentForm;
    private App_TodoList todoApp;
    private string recurringFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_recurring.txt");
    private Panel pnlMain, pnlSettings;
    private ListBox listTasks;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate, cmbDigest;
    private DateTimePicker dtpTime, dtpDigestTime;
    private Timer checkTimer;
    private string digestType = "不提醒", digestTimeStr = "08:00", lastDigestDate = "";
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    private class RecurringTask { public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate; }
    private List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        // 初始化面板
        pnlMain = new Panel() { Dock = DockStyle.Fill };
        pnlSettings = new Panel() { Dock = DockStyle.Fill, Visible = false, BackColor = Color.White };

        InitializeMainUI();
        InitializeSettingsUI();

        this.Controls.Add(pnlMain);
        this.Controls.Add(pnlSettings);

        LoadTasks();
        checkTimer = new Timer() { Interval = 600000, Enabled = true };
        checkTimer.Tick += (s, e) => CheckTasks();
    }

    private void InitializeMainUI() {
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 2 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75f));

        Label lblTitle = new Label() { Text = "週期任務清單", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5,0,0,0) };
        Button btnGoSet = new Button() { Text = "⚙ 設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,5,8) };
        btnGoSet.Click += (s, e) => { pnlMain.Visible = false; pnlSettings.Visible = true; };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnGoSet, 1, 0);
        pnlMain.Controls.Add(header);

        listTasks = new ListBox() { Dock = DockStyle.Fill, Font = new Font(MainFont.FontFamily, 10f), BorderStyle = BorderStyle.None };
        pnlMain.Controls.Add(listTasks);
        listTasks.BringToFront();

        Button btnRemove = new Button() { Text = "移除選取的排程", Dock = DockStyle.Bottom, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) };
        btnRemove.Click += (s, e) => { if(listTasks.SelectedIndex != -1) { tasks.RemoveAt(listTasks.SelectedIndex); SaveTasks(); RefreshUI(); } };
        pnlMain.Controls.Add(btnRemove);
    }

    private void InitializeSettingsUI() {
        FlowLayoutPanel setFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(15), AutoScroll = true };
        
        Button btnBack = new Button() { Text = "⬅ 返回清單", Width = 100, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = Color.Gray, ForeColor = Color.White };
        btnBack.Click += (s, e) => { pnlMain.Visible = true; pnlSettings.Visible = false; };
        setFlow.Controls.Add(btnBack);

        setFlow.Controls.Add(new Label() { Text = "【建立新排程】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) });

        setFlow.Controls.Add(new Label() { Text = "任務內容：", AutoSize = true });
        txtName = new TextBox() { Width = 300, Font = MainFont };
        setFlow.Controls.Add(txtName);

        FlowLayoutPanel timeRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 5, 0, 5) };
        cmbMonth = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbMonth.Items.Add("每個月"); for(int i=1;i<=12;i++) cmbMonth.Items.Add(i+"月"); cmbMonth.SelectedIndex = 0;
        
        cmbDate = new ComboBox() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDate.Items.Add("每天"); for(int i=1;i<=31;i++) cmbDate.Items.Add(i+"號");
        cmbDate.Items.AddRange(new string[]{"星期一","星期二","星期三","星期四","星期五","工作日","週末"}); cmbDate.SelectedIndex = 0;
        
        dtpTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        timeRow.Controls.AddRange(new Control[]{ cmbMonth, cmbDate, dtpTime });
        setFlow.Controls.Add(timeRow);

        Button btnAdd = new Button() { Text = "+ 建立週期任務", Width = 300, Height = 35, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold) };
        btnAdd.Click += (s, e) => AddTask();
        setFlow.Controls.Add(btnAdd);

        setFlow.Controls.Add(new Label() { Text = "【預告設定】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 20, 0, 5) });
        FlowLayoutPanel digestRow = new FlowLayoutPanel() { AutoSize = true };
        cmbDigest = new ComboBox() { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[]{"不提醒","每週一","每月1號"}); cmbDigest.SelectedIndex = 0;
        dtpDigestTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        Button btnSaveD = new Button() { Text = "儲存", Width = 60, Height = 25 };
        btnSaveD.Click += (s, e) => { digestType = cmbDigest.Text; digestTimeStr = dtpDigestTime.Value.ToString("HH:mm"); SaveTasks(); MessageBox.Show("預告設定已儲存"); };
        digestRow.Controls.AddRange(new Control[]{ cmbDigest, dtpDigestTime, btnSaveD });
        setFlow.Controls.Add(digestRow);

        pnlSettings.Controls.Add(setFlow);
    }

    // --- 核心邏輯 (與前版一致) ---
    private void AddTask() {
        if(string.IsNullOrEmpty(txtName.Text)) return;
        tasks.Add(new RecurringTask(){ Name=txtName.Text, MonthStr=cmbMonth.Text, DateStr=cmbDate.Text, TimeStr=dtpTime.Text, LastTriggeredDate="" });
        SaveTasks(); RefreshUI(); txtName.Text = ""; MessageBox.Show("排程建立完成！");
    }

    private void CheckTasks() {
        DateTime now = DateTime.Now; string today = now.ToString("yyyy-MM-dd"), time = now.ToString("HH:mm");
        bool needsSave = false;
        if(digestType != "不提醒" && lastDigestDate != today) {
            if((digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday) || (digestType == "每月1號" && now.Day == 1)) {
                if(string.Compare(time, digestTimeStr) >= 0) { lastDigestDate = today; needsSave = true; todoApp.AddTaskExternally("📅 [預告] 共有 " + tasks.Count + " 項週期任務"); }
            }
        }
        foreach(var t in tasks) {
            if(t.LastTriggeredDate == today) continue;
            bool mMatch = t.MonthStr == "每個月" || t.MonthStr == now.Month + "月";
            bool dMatch = t.DateStr == "每天" || t.DateStr == now.Day + "號" || (t.DateStr == "工作日" && now.DayOfWeek != DayOfWeek.Saturday && now.DayOfWeek != DayOfWeek.Sunday) || (t.DateStr == "星期" + "日一二三四五六"[(int)now.DayOfWeek]);
            if(mMatch && dMatch && string.Compare(time, t.TimeStr) >= 0) { t.LastTriggeredDate = today; needsSave = true; todoApp.AddTaskExternally(t.Name); }
        }
        if(needsSave) SaveTasks();
    }

    private void RefreshUI() { listTasks.Items.Clear(); foreach(var t in tasks) listTasks.Items.Add(string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name)); }
    private void SaveTasks() {
        List<string> lines = new List<string>(){ string.Format("#DIGEST|{0}|{1}|{2}", digestType, digestTimeStr, lastDigestDate) };
        foreach(var t in tasks) lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate));
        File.WriteAllLines(recurringFile, lines);
    }
    private void LoadTasks() {
        if(!File.Exists(recurringFile)) return;
        foreach(var l in File.ReadAllLines(recurringFile)) {
            if(l.StartsWith("#DIGEST")) { var p=l.Split('|'); digestType=p[1]; digestTimeStr=p[2]; lastDigestDate=p[3]; continue; }
            var p2=l.Split('|'); if(p2.Length>=5) tasks.Add(new RecurringTask(){ Name=p2[0], MonthStr=p2[1], DateStr=p2[2], TimeStr=p2[3], LastTriggeredDate=p2[4] });
        }
        RefreshUI();
    }
}
