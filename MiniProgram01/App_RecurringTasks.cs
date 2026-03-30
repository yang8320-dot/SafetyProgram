using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_RecurringTasks : UserControl {
    private MainForm parentForm;
    private App_TodoList todoApp;
    private string recurringFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_recurring.txt");
    private Panel pnlMain;
    private ListBox listTasks;
    private Timer checkTimer;

    // 全域設定變數 (供獨立視窗讀取)
    public string digestType = "不提醒";
    public string digestTimeStr = "08:00";
    public string lastDigestDate = "";
    public int advanceDays = 0;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    private class RecurringTask { public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate; }
    private List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; 
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        pnlMain = new Panel() { Dock = DockStyle.Fill };

        InitializeMainUI();
        this.Controls.Add(pnlMain);

        LoadTasks();
        checkTimer = new Timer() { Interval = 600000, Enabled = true };
        checkTimer.Tick += (s, e) => CheckTasks();
    }

    private void InitializeMainUI() {
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 2 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f)); // 加寬一點確保按鈕顯示完整

        Label lblTitle = new Label() { Text = "週期任務清單", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5,0,0,0) };
        
        // 點擊後跳出獨立視窗，並傳入目前的設定值
        Button btnGoSet = new Button() { Text = "⚙ 設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
        btnGoSet.Click += (s, e) => { 
            new RecurringSettingsWindow(this).ShowDialog(); 
        };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnGoSet, 1, 0);
        pnlMain.Controls.Add(header);

        listTasks = new ListBox() { Dock = DockStyle.Fill, Font = new Font(MainFont.FontFamily, 10f), BorderStyle = BorderStyle.None };
        pnlMain.Controls.Add(listTasks);
        listTasks.BringToFront();

        Button btnRemove = new Button() { Text = "移除選取的排程", Dock = DockStyle.Bottom, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand };
        btnRemove.Click += (s, e) => { if(listTasks.SelectedIndex != -1) { tasks.RemoveAt(listTasks.SelectedIndex); SaveTasks(); RefreshUI(); } };
        pnlMain.Controls.Add(btnRemove);
    }

    // ==========================================
    // 開放給設定視窗呼叫的 API
    // ==========================================
    public void AddNewRecurringTask(string name, string month, string date, string time) {
        tasks.Add(new RecurringTask() { Name = name, MonthStr = month, DateStr = date, TimeStr = time, LastTriggeredDate = "" });
        SaveTasks(); 
        RefreshUI(); 
        MessageBox.Show("排程建立完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void UpdateGlobalSettings(int advDays, string dType, string dTime) {
        this.advanceDays = advDays;
        this.digestType = dType;
        this.digestTimeStr = dTime;
        SaveTasks();
        MessageBox.Show("設定已儲存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ==========================================
    // 核心排程邏輯
    // ==========================================
    private void CheckTasks() {
        DateTime now = DateTime.Now; 
        string todayDate = now.ToString("yyyy-MM-dd"), nowTimeStr = now.ToString("HH:mm");
        bool needsSave = false;

        if(digestType != "不提醒" && lastDigestDate != todayDate) {
            bool isDigestDay = (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday) || (digestType == "每月1號" && now.Day == 1);
            if(isDigestDay && string.Compare(nowTimeStr, digestTimeStr) >= 0) { 
                lastDigestDate = todayDate; needsSave = true; 
                todoApp.AddTaskExternally("📅 [提醒] 共有 " + tasks.Count + " 項週期任務排程中"); 
            }
        }

        foreach(var t in tasks) {
            DateTime targetDate = now.AddDays(advanceDays);
            string targetKey = targetDate.ToString("yyyy-MM-dd");
            if(t.LastTriggeredDate == targetKey) continue;

            bool mMatch = t.MonthStr == "每個月" || t.MonthStr == "每週" || t.MonthStr == targetDate.Month + "月";
            DayOfWeek targetDOW = targetDate.DayOfWeek;
            bool dMatch = false;
            if (t.DateStr == "每天") dMatch = true;
            else if (t.DateStr == targetDate.Day + "號") dMatch = true;
            else if (t.DateStr == "工作日" && targetDOW != DayOfWeek.Saturday && targetDOW != DayOfWeek.Sunday) dMatch = true;
            else if (t.DateStr == "週末" && (targetDOW == DayOfWeek.Saturday || targetDOW == DayOfWeek.Sunday)) dMatch = true;
            else if (t.DateStr == "星期" + "日一二三四五六"[(int)targetDOW]) dMatch = true;

            if(mMatch && dMatch) {
                if(advanceDays > 0 || string.Compare(nowTimeStr, t.TimeStr) >= 0) {
                    t.LastTriggeredDate = targetKey; needsSave = true;
                    string prefix = advanceDays > 0 ? string.Format("[預排-{0}] ", targetDate.ToString("MM/dd")) : "";
                    todoApp.AddTaskExternally(prefix + t.Name);
                }
            }
        }
        if(needsSave) SaveTasks();
    }

    private void RefreshUI() { listTasks.Items.Clear(); foreach(var t in tasks) listTasks.Items.Add(string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name)); }
    
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
                if(p.Length >= 5) { int.TryParse(p[4], out advanceDays); }
                continue; 
            }
            var p2 = l.Split('|'); if(p2.Length >= 5) tasks.Add(new RecurringTask(){ Name=p2[0], MonthStr=p2[1], DateStr=p2[2], TimeStr=p2[3], LastTriggeredDate=p2[4] });
        }
        RefreshUI();
    }
}

// ==========================================
// 全新：獨立彈出的【⚙ 週期設定】視窗
// ==========================================
public class RecurringSettingsWindow : Form {
    private App_RecurringTasks parentControl;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate, cmbDigest, cmbAdvanceDays;
    private DateTimePicker dtpTime, dtpDigestTime;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public RecurringSettingsWindow(App_RecurringTasks parent) {
        this.parentControl = parent;
        this.Text = "⚙ 週期設定";
        this.Width = 380; 
        this.AutoSize = true;
        this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        this.StartPosition = FormStartPosition.CenterScreen; // 螢幕正中央
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        FlowLayoutPanel mainFlow = new FlowLayoutPanel() { 
            FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20), AutoSize = true 
        };

        // --- 框一：建立週期排程 ---
        GroupBox gbNew = new GroupBox() { Text = "建立週期排程", Font = new Font(MainFont, FontStyle.Bold), Width = 320, AutoSize = true, Margin = new Padding(0,0,0,15) };
        FlowLayoutPanel flowNew = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5, 15, 5, 10), AutoSize = true, Font = MainFont };
        
        flowNew.Controls.Add(new Label() { Text = "任務內容：", AutoSize = true, Margin = new Padding(0,0,0,5) });
        txtName = new TextBox() { Width = 290, Font = MainFont, Margin = new Padding(0,0,0,10) };
        flowNew.Controls.Add(txtName);

        FlowLayoutPanel timeRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        cmbMonth = new ComboBox() { Width = 75, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbMonth.Items.AddRange(new string[] { "每個月", "每週" });
        for(int i=1;i<=12;i++) cmbMonth.Items.Add(i+"月"); 
        cmbMonth.SelectedIndex = 0;
        
        cmbDate = new ComboBox() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDate.Items.Add("每天"); for(int i=1;i<=31;i++) cmbDate.Items.Add(i+"號");
        cmbDate.Items.AddRange(new string[]{"星期一","星期二","星期三","星期四","星期五","星期六","星期日","工作日","週末"}); cmbDate.SelectedIndex = 0;
        
        dtpTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        timeRow.Controls.AddRange(new Control[]{ cmbMonth, cmbDate, dtpTime });
        flowNew.Controls.Add(timeRow);

        Button btnAdd = new Button() { Text = "+ 建立週期任務", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 5, 0, 5), Cursor = Cursors.Hand };
        btnAdd.Click += (s, e) => {
            if (string.IsNullOrEmpty(txtName.Text.Trim())) return;
            parentControl.AddNewRecurringTask(txtName.Text.Trim(), cmbMonth.Text, cmbDate.Text, dtpTime.Text);
            txtName.Text = "";
        };
        flowNew.Controls.Add(btnAdd);
        gbNew.Controls.Add(flowNew);
        mainFlow.Controls.Add(gbNew);

        // --- 框二：全域排程設定 ---
        GroupBox gbSettings = new GroupBox() { Text = "排程與提醒設定", Font = new Font(MainFont, FontStyle.Bold), Width = 320, AutoSize = true };
        FlowLayoutPanel flowSettings = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5, 15, 5, 10), AutoSize = true, Font = MainFont };
        
        // 1. 自動新增設定
        FlowLayoutPanel advanceRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
        advanceRow.Controls.Add(new Label() { Text = "提前新增天數：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbAdvanceDays = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDown };
        cmbAdvanceDays.Items.AddRange(new string[] { "0", "1", "2", "3", "7" }); 
        cmbAdvanceDays.Text = parent.advanceDays.ToString();
        advanceRow.Controls.Add(cmbAdvanceDays);
        flowSettings.Controls.Add(advanceRow);

        // 2. 提醒設定
        FlowLayoutPanel digestRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        digestRow.Controls.Add(new Label() { Text = "清單提醒設定：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbDigest = new ComboBox() { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[]{"不提醒","每週一","每月1號"}); 
        cmbDigest.Text = parent.digestType;
        dtpDigestTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        DateTime dt; if(DateTime.TryParseExact(parent.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dt)) dtpDigestTime.Value = dt;
        digestRow.Controls.AddRange(new Control[]{ cmbDigest, dtpDigestTime });
        flowSettings.Controls.Add(digestRow);

        Button btnSaveD = new Button() { Text = "💾 儲存所有設定", Width = 290, Height = 40, FlatStyle = FlatStyle.Flat, BackColor = Color.WhiteSmoke, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand };
        btnSaveD.Click += (s, e) => { 
            int adv = 0; int.TryParse(cmbAdvanceDays.Text, out adv);
            parentControl.UpdateGlobalSettings(adv, cmbDigest.Text, dtpDigestTime.Value.ToString("HH:mm")); 
        };
        flowSettings.Controls.Add(btnSaveD);

        gbSettings.Controls.Add(flowSettings);
        mainFlow.Controls.Add(gbSettings);

        this.Controls.Add(mainFlow);
    }
}
