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
    private ComboBox cmbMonth; // 全新：月份下拉選單
    private ComboBox cmbDate;  // 全新：日期下拉選單
    private DateTimePicker dtpTime; // 時間選擇器
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
        public string MonthStr;  // 記錄月份 (每個月, 1月...)
        public string DateStr;   // 記錄日期 (每天, 1號, 星期一...)
        public string TimeStr; 
        public string LastTriggeredDate; 
    }
    
    private List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm;
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // ==========================================
        // 【全新介面排版】精算 X 座標，完美塞入三個選單
        // ==========================================
        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 155 }; 
        
        // 第一排：任務內容
        Label lblName = new Label() { Text = "任務內容：", Location = new Point(5, 12), AutoSize = true, Font = MainFont };
        txtName = new TextBox() { Location = new Point(85, 10), Width = 275, Font = MainFont, BorderStyle = BorderStyle.FixedSingle };

        // 第二排：月份、日期、時間 (三個選單)
        Label lblMonth = new Label() { Text = "月份：", Location = new Point(5, 47), AutoSize = true, Font = MainFont };
        cmbMonth = new ComboBox() { Location = new Point(48, 45), Width = 75, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbMonth.Items.Add("每個月");
        for (int i = 1; i <= 12; i++) cmbMonth.Items.Add(i + "月");
        cmbMonth.SelectedIndex = 0;

        Label lblDate = new Label() { Text = "日期：", Location = new Point(128, 47), AutoSize = true, Font = MainFont };
        cmbDate = new ComboBox() { Location = new Point(171, 45), Width = 80, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDate.Items.Add("每天");
        for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i + "號");
        cmbDate.Items.AddRange(new string[] { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "工作日", "週末" });
        cmbDate.SelectedIndex = 0;

        Label lblTime = new Label() { Text = "時間：", Location = new Point(256, 47), AutoSize = true, Font = MainFont };
        dtpTime = new DateTimePicker() { Location = new Point(299, 45), Width = 60, Font = MainFont, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };

        // 第三排：新增按鈕
        Button btnAdd = new Button() { 
            Text = "+ 建立週期任務", Location = new Point(85, 80), Width = 275, Height = 28, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand 
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += new EventHandler(delegate { AddRecurringTask(); });

        // 第四排：預告設定
        Label lblDigest = new Label() { Text = "總覽預告：", Location = new Point(5, 122), AutoSize = true, Font = MainFont, ForeColor = Color.DimGray };
        cmbDigest = new ComboBox() { Location = new Point(85, 120), Width = 95, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[] { "不提醒", "每週一", "每月1號" });
        cmbDigest.SelectedIndex = 0;

        dtpDigestTime = new DateTimePicker() { Location = new Point(185, 120), Width = 65, Font = MainFont, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };

        Button btnSaveDigest = new Button() { 
            Text = "儲存設定", Location = new Point(255, 119), Width = 105, Height = 26, 
            FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray, Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold)
        };
        btnSaveDigest.FlatAppearance.BorderSize = 0;
        btnSaveDigest.Click += delegate { 
            digestType = cmbDigest.SelectedItem.ToString();
            digestTimeStr = dtpDigestTime.Value.ToString("HH:mm");
            SaveTasks(); 
            MessageBox.Show("總覽預告設定已儲存！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
        };

        topPanel.Controls.Add(lblName); topPanel.Controls.Add(txtName);
        topPanel.Controls.Add(lblMonth); topPanel.Controls.Add(cmbMonth);
        topPanel.Controls.Add(lblDate); topPanel.Controls.Add(cmbDate);
        topPanel.Controls.Add(lblTime); topPanel.Controls.Add(dtpTime);
        topPanel.Controls.Add(btnAdd);
        
        topPanel.Controls.Add(lblDigest); topPanel.Controls.Add(cmbDigest);
        topPanel.Controls.Add(dtpDigestTime); topPanel.Controls.Add(btnSaveDigest);
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

        // 計時器：10 分鐘檢查一次
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

        // 【預告檢查】
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

        // ==========================================
        // 【全新智慧比對引擎】判斷「月份」與「日期」的組合
        // ==========================================
        foreach (RecurringTask t in tasks) {
            DateTime lastDate;
            bool hasLastDate = DateTime.TryParse(t.LastTriggeredDate, out lastDate);
            if (hasLastDate && lastDate.Date == now.Date) continue; // 今天發過了就跳過

            // 1. 檢查月份是否符合
            bool isMonthMatch = (t.MonthStr == "每個月") || (t.MonthStr == now.Month + "月");
            
            // 2. 檢查日期是否符合
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

            // 3. 綜合判定與時間檢查
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
                
                // 讀取總覽設定
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
                
                if (line.StartsWith("#")) continue; // 跳過註解

                string[] parts = line.Split('|');
                
                // 【無痛升級】如果讀到的是新的「5個欄位」格式
                if (parts.Length >= 5) {
                    tasks.Add(new RecurringTask() { Name = parts[0], MonthStr = parts[1], DateStr = parts[2], TimeStr = parts[3], LastTriggeredDate = parts[4] });
                } 
                // 【相容舊版】如果讀到的是舊的格式，自動幫你轉換成新的月份/日期！
                else if (parts.Length >= 3) {
                    string oldFreq = parts[1];
                    string newMonth = "每個月";
                    string newDate = "每天";
                    
                    if (oldFreq == "每週" || oldFreq == "星期一") newDate = "星期一";
                    else if (oldFreq == "工作日(一~五)") newDate = "工作日";
                    else if (oldFreq == "週末(六日)") newDate = "週末";
                    else if (oldFreq == "每月") newDate = "1號";
                    else if (oldFreq == "每季" || oldFreq == "每年") { newMonth = "1月"; newDate = "1號"; } // 舊版特例轉換
                    
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
            // 現在變成儲存 5 個欄位了
            lines.Add(string.Format("{0}|{1}|{2}|{3}|{4}", t.Name, t.MonthStr, t.DateStr, t.TimeStr, t.LastTriggeredDate));
        }
        File.WriteAllLines(recurringFile, lines.ToArray());
    }
}
