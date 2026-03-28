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
    private ComboBox cmbFreq;
    private DateTimePicker dtpTime;
    private ListBox listTasks;
    
    // 【預告系統專用 UI 與變數】
    private ComboBox cmbDigest;
    private DateTimePicker dtpDigestTime;
    private string digestType = "不提醒";
    private string digestTimeStr = "08:00";
    private string lastDigestDate = "";

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

        // --- 頂部設定區 (稍微加高以容納預告設定) ---
        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 135 };
        
        Label lblName = new Label() { Text = "任務內容：", Location = new Point(5, 8), AutoSize = true, Font = MainFont };
        txtName = new TextBox() { Location = new Point(80, 5), Width = 235, Font = MainFont, BorderStyle = BorderStyle.FixedSingle };

        Label lblFreq = new Label() { Text = "出現頻率：", Location = new Point(5, 40), AutoSize = true, Font = MainFont };
        cmbFreq = new ComboBox() { Location = new Point(80, 37), Width = 110, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
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

        // ==========================================
        // 【全新介面】總覽預告設定區
        // ==========================================
        Label lblDigest = new Label() { Text = "總覽預告：", Location = new Point(5, 107), AutoSize = true, Font = MainFont, ForeColor = Color.DimGray };
        cmbDigest = new ComboBox() { Location = new Point(80, 104), Width = 95, Font = MainFont, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[] { "不提醒", "每週一", "每月1號" });
        cmbDigest.SelectedIndex = 0;

        dtpDigestTime = new DateTimePicker() { Location = new Point(180, 104), Width = 65, Font = MainFont, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };

        Button btnSaveDigest = new Button() { 
            Text = "💾 儲存", Location = new Point(250, 103), Width = 65, Height = 26, 
            FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray, Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold)
        };
        btnSaveDigest.FlatAppearance.BorderSize = 0;
        btnSaveDigest.Click += delegate { 
            digestType = cmbDigest.SelectedItem.ToString();
            digestTimeStr = dtpDigestTime.Value.ToString("HH:mm");
            SaveTasks(); 
            MessageBox.Show("總覽預告設定已儲存！\n系統將在指定時間為您派發排程報告。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
        };

        topPanel.Controls.Add(lblName); topPanel.Controls.Add(txtName);
        topPanel.Controls.Add(lblFreq); topPanel.Controls.Add(cmbFreq);
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
        Button btnRemove = new Button() { Text = "🗑️ 移除選取的排程", Dock = DockStyle.Bottom, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) };
        btnRemove.FlatAppearance.BorderSize = 0;
        btnRemove.Click += new EventHandler(delegate { RemoveSelected(); });
        this.Controls.Add(btnRemove);

        LoadTasks();

        // 計時器：10 分鐘檢查一次
        checkTimer = new Timer();
        checkTimer.Interval = 600000; 
        checkTimer.Tick += new EventHandler(CheckTasksScheduler);
        checkTimer.Start();

        // 啟動時強制檢查一次
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
        MessageBox.Show("週期任務設定成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // ==========================================
        // 【核心邏輯 1】總覽預告檢查
        // ==========================================
        if (digestType != "不提醒" && lastDigestDate != todayStr) {
            bool triggerDigest = false;
            
            // 判斷是否符合預告條件
            if (digestType == "每週一" && dow == DayOfWeek.Monday && string.Compare(nowTimeStr, digestTimeStr) >= 0) {
                triggerDigest = true;
            }
            else if (digestType == "每月1號" && now.Day == 1 && string.Compare(nowTimeStr, digestTimeStr) >= 0) {
                triggerDigest = true;
            }

            if (triggerDigest) {
                lastDigestDate = todayStr;
                needsSave = true;

                // 打包所有的任務名稱
                List<string> upcoming = new List<string>();
                foreach(var tk in tasks) upcoming.Add(tk.Name);
                
                string summaryText = string.Join(", ", upcoming);
                // 避免字數太長把待辦清單撐破
                if (summaryText.Length > 40) summaryText = summaryText.Substring(0, 37) + "...";
                
                string title = digestType == "每週一" ? "📅 [本週預告]" : "📅 [本月預告]";
                string finalDigestTask = string.Format("{0} 共有 {1} 項排程準備執行 ({2})", title, tasks.Count, summaryText);
                
                todoApp.AddTaskExternally(finalDigestTask);
                parentForm.trayIcon.ShowBalloonTip(5000, "📅 排程預告", "已為您總整未來的排程，請至待辦清單查看！", ToolTipIcon.Info);
            }
        }

        // ==========================================
        // 【核心邏輯 2】獨立任務發派檢查
        // ==========================================
        foreach (RecurringTask t in tasks) {
            DateTime lastDate;
            bool hasLastDate = DateTime.TryParse(t.LastTriggeredDate, out lastDate);
            if (hasLastDate && lastDate.Date == now.Date) continue; 

            bool isTodayTheDay = false;
            switch (t.Frequency) {
                case "每天": isTodayTheDay = true; break;
                case "工作日(一~五)": isTodayTheDay = (dow >= DayOfWeek.Monday && dow <= DayOfWeek.Friday); break;
                case "週末(六日)": isTodayTheDay = (dow == DayOfWeek.Saturday || dow == DayOfWeek.Sunday); break;
                case "每週": isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddDays(7)); break;
                case "每月": isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddMonths(1)); break;
                case "每季": isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddMonths(3)); break;
                case "每年": isTodayTheDay = (!hasLastDate || now.Date >= lastDate.Date.AddYears(1)); break;
            }

            if (isTodayTheDay) {
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
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) {
                    // 讀取總覽設定 (隱藏儲存在第一行)
                    if (line.StartsWith("#DIGEST|")) {
                        string[] dParts = line.Split('|');
                        if (dParts.Length >= 4) {
                            cmbDigest.SelectedItem = dParts[1]; digestType = dParts[1];
                            DateTime pt; if (DateTime.TryParseExact(dParts[2], "HH:mm", null, System.Globalization.DateTimeStyles.None, out pt)) { dtpDigestTime.Value = pt; }
                            digestTimeStr = dParts[2];
                            lastDigestDate = dParts[3];
                        }
                    }
                    continue;
                }
                string[] parts = line.Split('|');
                if (parts.Length >= 3) {
                    tasks.Add(new RecurringTask() { Name = parts[0], Frequency = parts[1], TimeStr = parts[2], LastTriggeredDate = parts.Length > 3 ? parts[3] : "" });
                }
            }
        }
        RefreshListUI();
    }

    private void SaveTasks() {
        List<string> lines = new List<string>();
        // 悄悄將預告設定儲存在純文字檔的最上面
        lines.Add(string.Format("#DIGEST|{0}|{1}|{2}", digestType, digestTimeStr, lastDigestDate));
        lines.Add("# 格式說明：任務內容|出現頻率|時間|最後派發日期");
        foreach(RecurringTask t in tasks) {
            lines.Add(string.Format("{0}|{1}|{2}|{3}", t.Name, t.Frequency, t.TimeStr, t.LastTriggeredDate));
        }
        File.WriteAllLines(recurringFile, lines.ToArray());
    }
}
