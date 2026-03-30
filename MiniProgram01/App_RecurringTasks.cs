using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_RecurringTasks : UserControl {
    private MainForm parentForm;
    private App_TodoList todoApp;
    private string recurringFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_recurring.txt");
    private Timer checkTimer;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate, cmbDigest;
    private DateTimePicker dtpTime, dtpDigestTime;
    private ListBox listTasks;
    private string digestType = "不提醒", digestTimeStr = "08:00", lastDigestDate = "";

    private class RecurringTask { public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate; }
    private List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        FlowLayoutPanel top = new FlowLayoutPanel() { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        
        // 任務內容
        FlowLayoutPanel r1 = new FlowLayoutPanel() { AutoSize = true };
        r1.Controls.Add(new Label() { Text = "內容:", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        txtName = new TextBox() { Width = 260 };
        r1.Controls.Add(txtName);

        // 時間設定
        FlowLayoutPanel r2 = new FlowLayoutPanel() { AutoSize = true };
        cmbMonth = new ComboBox() { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbMonth.Items.Add("每個月"); for(int i=1;i<=12;i++) cmbMonth.Items.Add(i+"月"); cmbMonth.SelectedIndex = 0;
        cmbDate = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDate.Items.Add("每天"); for(int i=1;i<=31;i++) cmbDate.Items.Add(i+"號");
        cmbDate.Items.AddRange(new string[]{"星期一","星期二","星期三","星期四","星期五","工作日","週末"}); cmbDate.SelectedIndex = 0;
        dtpTime = new DateTimePicker() { Width = 65, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        r2.Controls.AddRange(new Control[]{ cmbMonth, cmbDate, dtpTime });

        Button btnAdd = new Button() { Text = "+ 建立週期任務", Width = 310, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
        btnAdd.Click += (s, e) => AddTask();

        // 預告區
        FlowLayoutPanel r4 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
        r4.Controls.Add(new Label() { Text = "預告:", AutoSize = true });
        cmbDigest = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[]{"不提醒","每週一","每月1號"}); cmbDigest.SelectedIndex = 0;
        dtpDigestTime = new DateTimePicker() { Width = 65, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        Button btnSaveD = new Button() { Text = "儲存", Width = 50 };
        btnSaveD.Click += (s, e) => { digestType = cmbDigest.Text; digestTimeStr = dtpDigestTime.Value.ToString("HH:mm"); SaveTasks(); MessageBox.Show("預告已儲存"); };
        r4.Controls.AddRange(new Control[]{ cmbDigest, dtpDigestTime, btnSaveD });

        top.Controls.AddRange(new Control[]{ r1, r2, btnAdd, r4 });
        this.Controls.Add(top);

        listTasks = new ListBox() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None };
        this.Controls.Add(listTasks); listTasks.BringToFront();

        Button btnDel = new Button() { Text = "移除排程", Dock = DockStyle.Bottom, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White };
        btnDel.Click += (s, e) => { if(listTasks.SelectedIndex!=-1){ tasks.RemoveAt(listTasks.SelectedIndex); SaveTasks(); RefreshUI(); } };
        this.Controls.Add(btnDel);

        LoadTasks();
        checkTimer = new Timer() { Interval = 600000, Enabled = true };
        checkTimer.Tick += (s, e) => CheckEngine();
        CheckEngine();
    }

    private void AddTask() {
        if (string.IsNullOrEmpty(txtName.Text)) return;
        tasks.Add(new RecurringTask() { Name=txtName.Text, MonthStr=cmbMonth.Text, DateStr=cmbDate.Text, TimeStr=dtpTime.Text, LastTriggeredDate="" });
        SaveTasks(); RefreshUI(); txtName.Text = "";
    }

    private void CheckEngine() {
        DateTime now = DateTime.Now; string today = now.ToString("yyyy-MM-dd"), time = now.ToString("HH:mm");
        bool needsSave = false;

        // 預告檢查
        if (digestType != "不提醒" && lastDigestDate != today) {
            bool hit = (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday) || (digestType == "每月1號" && now.Day == 1);
            if (hit && string.Compare(time, digestTimeStr) >= 0) {
                lastDigestDate = today; needsSave = true;
                todoApp.AddTaskExternally(string.Format("📅 [預告] 共有 {0} 項週期排程", tasks.Count));
            }
        }

        // 任務檢查
        foreach (var t in tasks) {
            if (t.LastTriggeredDate == today) continue;
            bool mMatch = t.MonthStr == "每個月" || t.MonthStr == now.Month + "月";
            bool dMatch = t.DateStr == "每天" || t.DateStr == now.Day + "號" || 
                         (t.DateStr == "工作日" && now.DayOfWeek != DayOfWeek.Saturday && now.DayOfWeek != DayOfWeek.Sunday) ||
                         (t.DateStr == "星期" + "日一二三四五六"[(int)now.DayOfWeek].ToString());
            if (mMatch && dMatch && string.Compare(time, t.TimeStr) >= 0) {
                t.LastTriggeredDate = today; needsSave = true;
                todoApp.AddTaskExternally(t.Name);
            }
        }
        if (needsSave) SaveTasks();
    }

    private void RefreshUI() { listTasks.Items.Clear(); foreach(var t in tasks) listTasks.Items.Add(string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name)); }
    private void SaveTasks() {
        List<string> lines = new List<string>() { string.Format("#DIGEST|{0}|{1}|{2}", digestType, digestTimeStr, lastDigestDate) };
        foreach(var t in tasks) lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate));
        File.WriteAllLines(recurringFile, lines);
    }
    private void LoadTasks() {
        if(!File.Exists(recurringFile)) return;
        foreach(var l in File.ReadAllLines(recurringFile)){
            if(l.StartsWith("#DIGEST")){ var p=l.Split('|'); digestType=p[1]; digestTimeStr=p[2]; lastDigestDate=p[3]; continue; }
            var p2=l.Split('|'); if(p2.Length>=5) tasks.Add(new RecurringTask(){ Name=p2[0], MonthStr=p2[1], DateStr=p2[2], TimeStr=p2[3], LastTriggeredDate=p2[4] });
        }
        RefreshUI();
    }
}
