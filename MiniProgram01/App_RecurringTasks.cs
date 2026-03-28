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
    
    // UI 元件
    private TextBox txtName;
    private ComboBox cmbMonth; 
    private ComboBox cmbDate;  
    private DateTimePicker dtpTime; 
    private ListBox listTasks;
    
    // 預告系統專用 UI 與變數
    private ComboBox cmbDigest;
    private DateTimePicker dtpDigestTime;
    private string digestType = "不提醒";
    private string digestTimeStr = "08:00";
    private string lastDigestDate = "";

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    private class RecurringTask {
        public string Name;
        public string MonthStr;  
        public string DateStr;   
        public string TimeStr; 
        public string LastTriggeredDate; 
    }
    
    private List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm;
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(8);

        // ==========================================
        // 【無敵流式排版】棄用絕對座標，改用 FlowLayoutPanel
        // ==========================================
        
        // 主容器：由上往下排列，高度會隨內容自動長高
        FlowLayoutPanel topPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, 
            FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(0, 0, 0, 10) 
        }; 

        // --- 第 1 行：任務內容 ---
        FlowLayoutPanel r1 = new FlowLayoutPanel() { AutoSize = true, WrapContents = true, Margin = new Padding(0, 5, 0, 5) };
        Label lblName = new Label() { Text = "任務內容：", AutoSize = true, Font = MainFont, Margin = new Padding(3, 6, 3, 0) };
        txtName = new TextBox() { Width = 230, Font = MainFont };
        r1.Controls.Add(lblName); r1.Controls.Add(txtName);

        // --- 第 2 行：排程時間 (月份 + 日期 + 時間) ---
        FlowLayoutPanel r2 = new FlowLayoutPanel() { AutoSize = true, WrapContents = true, Margin = new Padding(0, 5, 0, 5) };
        Label lblFreq = new Label() { Text = "排程時間：", AutoSize = true, Font = MainFont, Margin = new Padding(3, 6, 3, 0) };
        
        cmbMonth = new ComboBox() { Width = 75, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbMonth.Items.Add("每個月");
        for (int i = 1; i <= 12; i++) cmbMonth.Items.Add(i + "月");
        cmbMonth.SelectedIndex = 0;

        cmbDate = new ComboBox() { Width = 85, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDate.Items.Add("每天");
        for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i + "號");
        cmbDate.Items.AddRange(new string[] { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "工作日", "週末" });
        cmbDate.SelectedIndex = 0;

        dtpTime = new DateTimePicker() { Width = 65, Font = MainFont, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        
        r2.Controls.Add(lblFreq); r2.Controls.Add(cmbMonth); r2.Controls.Add(cmbDate); r2.Controls.Add(dtpTime);

        // --- 第 3 行：大大的建立按鈕 ---
        FlowLayoutPanel r3 = new FlowLayoutPanel() { AutoSize = true, WrapContents = true, Margin = new Padding(0, 5, 0, 5) };
        Button btnAdd = new Button() { 
            Text = "+ 建立週期任務", Width = 310, Height = 32, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand,
            Margin = new Padding(5, 0, 0, 0)
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += new EventHandler(delegate { AddRecurringTask(); });
        r3.Controls.Add(btnAdd);

        // --- 第 4 行：總覽預告 ---
        FlowLayoutPanel r4 = new FlowLayoutPanel() { AutoSize = true, WrapContents = true, Margin = new Padding(0, 15, 0, 5) };
        Label lblDigest = new Label() { Text = "總覽預告：", AutoSize = true, Font = MainFont, ForeColor = Color.DimGray, Margin = new Padding(3, 6, 3, 0) };
        
        cmbDigest = new ComboBox() { Width = 85, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[] { "不提醒", "每週一", "每月1號" });
        cmbDigest.SelectedIndex = 0;

        dtpDigestTime = new DateTimePicker() { Width = 65, Font = MainFont, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };

        Button btnSaveDigest = new Button() { 
            Text = "儲存", Width = 60, Height = 28, 
            FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray, Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold)
        };
        btnSaveDigest.FlatAppearance.BorderSize = 0;
        btnSaveDigest.Click += delegate { 
            digestType = cmbDigest.SelectedItem.ToString();
            digestTimeStr = dtpDigestTime.Value.ToString("HH:mm");
            SaveTasks(); 
            MessageBox.Show("總覽預告設定已儲存！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
        };

        r4.Controls.Add(lblDigest); r4.Controls.Add(cmbDigest); r4.Controls.Add(dtpDigestTime); r4.Controls.Add(btnSaveDigest);

        // 將這四行塞入主容器
        topPanel.Controls.Add(r1);
        topPanel.Controls.Add(r2);
        topPanel.Controls.Add(r3);
        topPanel.Controls.Add(r4);
        
        this.Controls.Add(topPanel);

        // --- 任務清單顯示區 ---
        listTasks = new ListBox() { Dock = DockStyle.Fill, Font = new Font(MainFont.FontFamily, 10f), BorderStyle = BorderStyle.None };
        this.Controls.Add(listTasks);
        listTasks.BringToFront();

        // --- 底部刪除按鈕 ---
        Button btnRemove = new Button() { Text = "移除選取的排程", Dock = DockStyle.Bottom, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) };
        btnRemove.FlatAppearance.BorderSize = 0;
        btnRemove.Click += new EventHandler(delegate { RemoveSelected(); });
        this.Controls.Add(btnRemove);

        LoadTasks();

        checkTimer = new Timer();
        checkTimer.Interval = 600000; 
        checkTimer.Tick += new EventHandler(CheckTasksScheduler);
        checkTimer.Start();

        CheckTasksScheduler(null, null);
    }

    private void AddRecurringTask() {
        string name = txtName.Text.Trim();
        if (string.IsNullOrEmpty(name)) return;
        
        string month = cmbMonth.SelectedItem.ToString();
        string date = cmbDate.SelectedItem.ToString();
        string time = dtpTime.Value.ToString("HH:mm");

        tasks.Add(new RecurringTask() { Name = name, MonthStr = month, DateStr = date, TimeStr = time, LastTriggeredDate = "" });
        SaveTasks();
        RefreshListUI();
        
        txtName.Text = "";
        MessageBox.Show("設定成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void RemoveSelected() {
        if (listTasks.SelectedIndex != -1) {
            tasks.RemoveAt(listTasks.SelectedIndex);
            SaveTasks();
            RefreshListUI();
        }
    }

    private void CheckTasksScheduler(object sender, EventArgs e) {
        bool needsSave = false;
        DateTime now = DateTime.Now;
        string todayStr = now.ToString("yyyy-MM-dd");
        string nowTimeStr = now.ToString("HH:mm");
        DayOfWeek dow = now.DayOfWeek;

        if (digestType != "不提醒" && lastDigestDate != todayStr) {
            bool triggerDigest = false;
            if (digestType == "每週一" && dow == DayOfWeek.Monday && string.Compare(nowTimeStr, digestTimeStr) >= 0) triggerDigest = true;
            else if (digestType == "每月1號" && now.Day == 1 && string.Compare(nowTimeStr, digestTimeStr) >= 0) triggerDigest = true;

            if (triggerDigest) {
                lastDigestDate = todayStr; needsSave = true;
                List<string> upcoming = new List<string>();
                foreach(var tk in tasks) upcoming.Add(tk.Name);
                string summaryText = string.Join(", ", upcoming);
                if (summaryText.Length > 40) summaryText = summaryText.Substring(0, 37) + "...";
                
                string title = digestType == "每週一" ? "📅 [本週預告]" : "📅 [本月預告]";
                todoApp.AddTaskExternally(string.Format("{0} 共有 {1} 項排程準備執行 ({2})", title, tasks.Count, summaryText));
                parentForm.trayIcon.ShowBalloonTip(5000, "📅 排程預告", "已為您總整未來的排程，請至待辦清單查看！", ToolTipIcon.Info);
            }
        }

        foreach (RecurringTask t in tasks) {
            DateTime lastDate;
            bool hasLastDate = DateTime.TryParse(t.LastTriggeredDate, out lastDate);
            if (hasLastDate && lastDate.Date == now.Date) continue; 

            bool isMonthMatch = (t.MonthStr == "每個月") || (t.MonthStr == now.Month + "月");
            bool isDateMatch = false;
            
            if (t.DateStr == "每天") isDateMatch = true;
            else if (t.DateStr == now.Day + "號") isDateMatch = true;
            else if (t.DateStr == "工作日" && dow >= DayOfWeek.Monday && dow <= DayOfWeek.Friday) isDateMatch = true;
            else if (t.DateStr == "週末" && (dow == DayOfWeek.Saturday || dow == DayOfWeek.Sunday)) isDateMatch = true;
            else {
                switch(t.DateStr) {
                    case "星期一": isDateMatch = (dow == DayOfWeek.Monday); break;
                    case "星期二": isDateMatch = (dow == DayOfWeek.Tuesday); break;
                    case "星期三": isDateMatch = (dow == DayOfWeek.Wednesday); break;
                    case "星期四": isDateMatch = (dow == DayOfWeek.Thursday); break;
                    case "星期五": isDateMatch = (dow == DayOfWeek.Friday); break;
                    case "星期六": isDateMatch = (dow == DayOfWeek.Saturday); break;
                    case "星期日": isDateMatch = (dow == DayOfWeek.Sunday); break;
                }
            }

            if (isMonthMatch && isDateMatch) {
                if (string.Compare(nowTimeStr, t.TimeStr) >= 0) {
                    t.LastTriggeredDate = todayStr; 
                    needsSave = true;
                    todoApp.AddTaskExternally(t.Name);
                    parentForm.trayIcon.ShowBalloonTip(5000, "⏰ 週期任務", "新增排程：" + t.Name, ToolTipIcon.Info);
                }
            }
        }

        if (needsSave) SaveTasks();
    }

    private void RefreshListUI() {
        listTasks.Items.Clear();
        foreach (RecurringTask t in tasks) {
            listTasks.Items.Add(string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name));
        }
    }

    private void LoadTasks() {
        tasks.Clear();
        if (File.Exists(recurringFile)) {
            string[] lines = File.ReadAllLines(recurringFile);
            foreach(string line in lines) {
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (line.StartsWith("#DIGEST|")) {
                    string[] dParts = line.Split('|');
                    if (dParts.Length >= 4) {
                        cmbDigest.SelectedItem = dParts[1]; digestType = dParts[1];
                        DateTime pt; if (DateTime.TryParseExact(dParts[2], "HH:mm", null, System.Globalization.DateTimeStyles.None, out pt)) { dtpDigestTime.Value = pt; }
                        digestTimeStr = dParts[2];
                        lastDigestDate = dParts[3];
                    }
                    continue;
                }
                
                if (line.StartsWith("#")) continue; 

                string[] parts = line.Split('|');
                if (parts.Length >= 5) {
                    tasks.Add(new RecurringTask() { Name = parts[0], MonthStr = parts[1], DateStr = parts[2], TimeStr = parts[3], LastTriggeredDate = parts[4] });
                } 
                else if (parts.Length >= 3) {
                    string oldFreq = parts[1];
                    string newMonth = "每個月";
                    string newDate = "每天";
                    
                    if (oldFreq == "每週" || oldFreq == "星期一") newDate = "星期一";
                    else if (oldFreq == "工作日(一~五)") newDate = "工作日";
                    else if (oldFreq == "週末(六日)") newDate = "週末";
                    else if (oldFreq == "每月") newDate = "1號";
                    else if (oldFreq == "每季" || oldFreq == "每年") { newMonth = "1月"; newDate = "1號"; }
                    
                    tasks.Add(new RecurringTask() { Name = parts[0], MonthStr = newMonth, DateStr = newDate, TimeStr = parts[2], LastTriggeredDate = parts.Length > 3 ? parts[3] : "" });
                }
            }
        }
        RefreshListUI();
    }

    private void SaveTasks() {
        List<string> lines = new List<string>();
        lines.Add(string.Format("#DIGEST|{0}|{1}|{2}", digestType, digestTimeStr, lastDigestDate));
        lines.Add("# 格式說明：任務內容|月份|日期|時間|最後派發日期");
        foreach(RecurringTask t in tasks) {
            lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate));
        }
        File.WriteAllLines(recurringFile, lines.ToArray());
    }
}
