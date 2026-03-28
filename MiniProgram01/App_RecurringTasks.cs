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
    private ComboBox cmbFreq;
    private DateTimePicker dtpTime;
    private ListBox listTasks;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    private class RecurringTask {
        public string Name;
        public string Frequency; 
        public string TimeStr; 
        public string LastTriggeredDate; 
    }
    
    private List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm;
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // --- 頂部設定區 ---
        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 100 };
        
        Label lblName = new Label() { Text = "任務內容：", Location = new Point(5, 8), AutoSize = true, Font = MainFont };
        txtName = new TextBox() { Location = new Point(80, 5), Width = 235, Font = MainFont, BorderStyle = BorderStyle.FixedSingle };

        Label lblFreq = new Label() { Text = "出現頻率：", Location = new Point(5, 40), AutoSize = true, Font = MainFont };
        cmbFreq = new ComboBox() { Location = new Point(80, 37), Width = 110, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        
        // 【全新升級】加入日、週、月、季、年，並保留實用的工作日選項
        cmbFreq.Items.AddRange(new string[] { "每天", "每週", "每月", "每季", "每年", "工作日(一~五)", "週末(六日)" });
        cmbFreq.SelectedIndex = 0;

        Label lblTime = new Label() { Text = "時間：", Location = new Point(195, 40), AutoSize = true, Font = MainFont };
        dtpTime = new DateTimePicker() { Location = new Point(245, 37), Width = 70, Font = MainFont, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };

        Button btnAdd = new Button() { 
            Text = "➕ 建立週期任務", Location = new Point(80, 68), Width = 235, Height = 28, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand 
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += new EventHandler(delegate { AddRecurringTask(); });

        topPanel.Controls.Add(lblName); topPanel.Controls.Add(txtName);
        topPanel.Controls.Add(lblFreq); topPanel.Controls.Add(cmbFreq);
        topPanel.Controls.Add(lblTime); topPanel.Controls.Add(dtpTime);
        topPanel.Controls.Add(btnAdd);
        this.Controls.Add(topPanel);

        // --- 任務清單顯示區 ---
        listTasks = new ListBox() { Dock = DockStyle.Fill, Font = new Font(MainFont.FontFamily, 10f), BorderStyle = BorderStyle.None };
        this.Controls.Add(listTasks);
        listTasks.BringToFront();

        // --- 底部刪除按鈕 ---
        Button btnRemove = new Button() { Text = "🗑️ 移除選取的排程", Dock = DockStyle.Bottom, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) };
        btnRemove.FlatAppearance.BorderSize = 0;
        btnRemove.Click += new EventHandler(delegate { RemoveSelected(); });
        this.Controls.Add(btnRemove);

        LoadTasks();

        // ==========================================
        // 【核心修改】將計時器改為 10 分鐘檢查一次
        // 10 分鐘 = 10 * 60 秒 * 1000 毫秒 = 600,000 毫秒
        // ==========================================
        checkTimer = new Timer();
        checkTimer.Interval = 600000; 
        checkTimer.Tick += new EventHandler(CheckTasksScheduler);
        checkTimer.Start();

        // 程式剛打開時，立刻強制檢查一次
        CheckTasksScheduler(null, null);
    }

    private void AddRecurringTask() {
        string name = txtName.Text.Trim();
        if (string.IsNullOrEmpty(name)) return;

        string freq = cmbFreq.SelectedItem.ToString();
        string time = dtpTime.Value.ToString("HH:mm");

        tasks.Add(new RecurringTask() { Name = name, Frequency = freq, TimeStr = time, LastTriggeredDate = "" });
        SaveTasks();
        RefreshListUI();
        
        txtName.Text = "";
        MessageBox.Show("設定成功！時間到時將自動派發至待辦清單。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void RemoveSelected() {
        if (listTasks.SelectedIndex != -1) {
            tasks.RemoveAt(listTasks.SelectedIndex);
            SaveTasks();
            RefreshListUI();
        }
    }

    // ==========================================
    // 【核心邏輯】支援日/週/月/季/年的智慧時間推算
    // ==========================================
    private void CheckTasksScheduler(object sender, EventArgs e) {
        bool needsSave = false;
        DateTime now = DateTime.Now;
        string todayStr = now.ToString("yyyy-MM-dd");
        string nowTimeStr = now.ToString("HH:mm");
        DayOfWeek dow = now.DayOfWeek;

        foreach (RecurringTask t in tasks) {
            DateTime lastDate;
            bool hasLastDate = DateTime.TryParse(t.LastTriggeredDate, out lastDate);

            // 防呆：如果今天已經派發過，絕對略過（防止同一天內重複發送）
            if (hasLastDate && lastDate.Date == now.Date) continue; 

            bool isTodayTheDay = false;

            switch (t.Frequency) {
                case "每天": 
                    isTodayTheDay = true; 
                    break;
                case "工作日(一~五)": 
                    isTodayTheDay = (dow >= DayOfWeek.Monday && dow <= DayOfWeek.Friday); 
                    break;
                case "週末(六日)": 
                    isTodayTheDay = (dow == DayOfWeek.Saturday || dow == DayOfWeek.Sunday); 
                    break;
                case "每週": 
                    // 如果沒派發過，今天就是第一天；如果有派發過，檢查是否過了 7 天
                    isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddDays(7)); 
                    break;
                case "每月": 
                    // 檢查是否過了一個月 (例如 3/28 -> 4/28)
                    isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddMonths(1)); 
                    break;
                case "每季": 
                    // 檢查是否過了三個月
                    isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddMonths(3)); 
                    break;
                case "每年": 
                    // 檢查是否過了一年
                    isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddYears(1)); 
                    break;
                // 為了相容如果你之前有用舊代碼存過特定的星期幾
                case "每週一": isTodayTheDay = (dow == DayOfWeek.Monday); break;
                case "每週二": isTodayTheDay = (dow == DayOfWeek.Tuesday); break;
                case "每週三": isTodayTheDay = (dow == DayOfWeek.Wednesday); break;
                case "每週四": isTodayTheDay = (dow == DayOfWeek.Thursday); break;
                case "每週五": isTodayTheDay = (dow == DayOfWeek.Friday); break;
            }

            if (isTodayTheDay) {
                // 時間比對：現在時間 >= 設定時間
                if (string.Compare(nowTimeStr, t.TimeStr) >= 0) {
                    t.LastTriggeredDate = todayStr; 
                    needsSave = true;
                    todoApp.AddTaskExternally(t.Name);
                    parentForm.trayIcon.ShowBalloonTip(5000, "⏰ 週期任務提醒", "已自動新增待辦事項：\n" + t.Name, ToolTipIcon.Info);
                }
            }
        }

        if (needsSave) SaveTasks();
    }

    private void RefreshListUI() {
        listTasks.Items.Clear();
        foreach (RecurringTask t in tasks) {
            listTasks.Items.Add(string.Format("[{0} {1}] {2}", t.Frequency, t.TimeStr, t.Name));
        }
    }

    private void LoadTasks() {
        tasks.Clear();
        if (File.Exists(recurringFile)) {
            string[] lines = File.ReadAllLines(recurringFile);
            foreach(string line in lines) {
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;
                string[] parts = line.Split('|');
                if (parts.Length >= 3) {
                    tasks.Add(new RecurringTask() { 
                        Name = parts[0], 
                        Frequency = parts[1], 
                        TimeStr = parts[2], 
                        LastTriggeredDate = parts.Length > 3 ? parts[3] : "" 
                    });
                }
            }
        }
        RefreshListUI();
    }

    private void SaveTasks() {
        List<string> lines = new List<string>();
        lines.Add("# 格式說明：任務內容|出現頻率|時間|最後派發日期");
        foreach(RecurringTask t in tasks) {
            lines.Add(string.Format("{0}|{1}|{2}|{3}", t.Name, t.Frequency, t.TimeStr, t.LastTriggeredDate));
        }
        File.WriteAllLines(recurringFile, lines.ToArray());
    }
}
