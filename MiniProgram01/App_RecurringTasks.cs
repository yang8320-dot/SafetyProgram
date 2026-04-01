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

    // 全域設定變數
    public string digestType { get; set; } = "不提醒";
    public string digestTimeStr { get; set; } = "08:00";
    public string lastDigestDate { get; set; } = "";
    public int advanceDays { get; set; } = 0;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public class RecurringTask { 
        public string Name, MonthStr, DateStr, TimeStr, LastTriggeredDate; 
    }
    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; 
        this.todoApp = todoApp;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        // --- 頂部標題列 ---
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 4 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label() { Text = "週期任務", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5,0,0,0) };
        
        Button btnViewAll = new Button() { Text = "全部檢視", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke };
        btnViewAll.Click += (s, e) => OpenAllTasksView();

        Button btnAdd = new Button() { Text = "新增任務", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand, BackColor = Color.White };
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this).ShowDialog(); };
        
        Button btnSet = new Button() { Text = "排程設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        header.Controls.AddRange(new Control[] { lblTitle, btnViewAll, btnAdd, btnSet });
        this.Controls.Add(header);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        
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
        checkTimer = new Timer() { Interval = 600000, Enabled = true };
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks();
    }

    public void OpenAllTasksView() {
        this.Invoke(new Action(() => { new AllTasksViewWindow(this).Show(); }));
    }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 25 : 450;

        foreach (var t in tasks) {
            Panel card = new Panel() { 
                Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 45), 
                BorderStyle = BorderStyle.None, Margin = new Padding(5, 5, 5, 8), 
                BackColor = Color.FromArgb(248, 248, 250) 
            };
            
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, AutoSize = true, Padding = new Padding(5) };
            
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button btnEdit = new Button() { Text = "調", Dock = DockStyle.Fill, Height = 28, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0,0,3,0) };
            btnEdit.FlatAppearance.BorderSize = 0; 
            btnEdit.Click += (s, e) => { 
                int currentIdx = tasks.IndexOf(t);
                if(currentIdx != -1) { new EditRecurringTaskWindow(this, currentIdx, t).ShowDialog(); RefreshUI(); }
            };

            Button btnDel = new Button() { Text = "✕", Dock = DockStyle.Fill, Height = 28, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0,0,3,0) };
            btnDel.FlatAppearance.BorderSize = 0; 
            btnDel.Click += (s, e) => { 
                if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { 
                    DeleteTask(t); 
                } 
            };

            Label lbl = new Label() { Text = string.Format("[{0} {1} {2}] {3}", t.MonthStr, t.DateStr, t.TimeStr, t.Name), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Font = MainFont };
            
            tlp.Controls.Add(btnEdit, 0, 0); 
            tlp.Controls.Add(btnDel, 1, 0); 
            tlp.Controls.Add(lbl, 2, 0);
            
            card.Controls.Add(tlp); 
            taskPanel.Controls.Add(card);
        }
    }

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

    public void DeleteTask(RecurringTask task) {
        if (tasks.Contains(task)) {
            tasks.Remove(task);
            SaveTasks();
            RefreshUI();
        }
    }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays) {
        digestType = dType; digestTimeStr = dTime; advanceDays = aDays;
        SaveTasks(); MessageBox.Show("設定儲存成功！");
    }

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
                        todoApp.AddTask(prefix + t.Name); 
                        t.LastTriggeredDate = targetDateStr; needsSave = true;
                        parentForm.AlertTab(1);
                    }
                }
            }
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
                    lastDigestDate = todayStr; needsSave = true;
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
            int h = int.Parse(timeParts[0]); int m = int.Parse(timeParts[1]);
            if (t.MonthStr == "每天") {
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                if (now > target) target = target.AddDays(1);
                return true;
            } else if (t.MonthStr == "每週") {
                Dictionary<string, DayOfWeek> dow = new Dictionary<string, DayOfWeek>() { {"一", DayOfWeek.Monday}, {"二", DayOfWeek.Tuesday}, {"三", DayOfWeek.Wednesday}, {"四", DayOfWeek.Thursday}, {"五", DayOfWeek.Friday}, {"六", DayOfWeek.Saturday}, {"日", DayOfWeek.Sunday} };
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
            } else if (t.MonthStr.EndsWith("月")) { 
                int month = int.Parse(t.MonthStr.Replace("月",""));
                int day = (t.DateStr == "月底") ? DateTime.DaysInMonth(now.Year, month) : int.Parse(t.DateStr);
                target = new DateTime(now.Year, month, day, h, m, 0);
                if (now > target) target = target.AddYears(1);
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
                var p = l.Split('|'); digestType = p[1]; digestTimeStr = p[2]; lastDigestDate = p[3]; 
                if(p.Length >= 5) { int adv = 0; int.TryParse(p[4], out adv); advanceDays = adv; }
            } else {
                var p = l.Split('|');
                if(p.Length >= 4) tasks.Add(new RecurringTask() { Name = p[0], MonthStr = p[1], DateStr = p[2], TimeStr = p[3], LastTriggeredDate = p.Length > 4 ? p[4] : "" });
            }
        }
        RefreshUI();
    }
}

// ==========================================
// 檢視全部任務的子視窗
// ==========================================
public class AllTasksViewWindow : Form {
    private App_RecurringTasks parentControl;
    private FlowLayoutPanel flow;
    
    // 分頁控制變數
    private int currentCategoryIndex = 0;
    private int currentItemIndex = 0;
    private int pageNum = 1;

    public AllTasksViewWindow(App_RecurringTasks parent) {
        this.parentControl = parent;
        this.Text = "全部排程清單總覽";
        
        Rectangle screenArea = Screen.FromPoint(Cursor.Position).WorkingArea;
        this.Width = 800; 
        this.Height = Math.Min(950, screenArea.Height - 120); 
        this.StartPosition = FormStartPosition.Manual;
        this.Left = screenArea.Left + (screenArea.Width - this.Width) / 2; 
        this.Top = screenArea.Top + 50; 
        
        this.BackColor = Color.White;
        this.Font = new Font("Microsoft JhengHei UI", 10.5f);

        Panel headerArea = new Panel() { Dock = DockStyle.Top, Height = 60, BackColor = Color.WhiteSmoke };
        Label lblTitle = new Label() { Text = "週期任務排程總覽 (全紀錄)", Left = 20, Top = 15, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 16f, FontStyle.Bold) };
        
        Button btnExport = new Button() { 
            Text = "轉存 PDF / 列印", Left = 630, Top = 12, Width = 150, Height = 35, 
            BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand 
        };
        btnExport.Click += (s, e) => ExportToPDF();

        headerArea.Controls.AddRange(new Control[] { lblTitle, btnExport });
        this.Controls.Add(headerArea);

        flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        this.Controls.Add(flow);
        RefreshData();
    }

    public void RefreshData() {
        flow.Controls.Clear();
        var tasks = parentControl.tasks;
        AddGroup(flow, "【 每天觸發 】", tasks.Where(t => t.MonthStr == "每天").ToList());
        AddGroup(flow, "【 每週觸發 】", tasks.Where(t => t.MonthStr == "每週").ToList());
        AddGroup(flow, "【 每月觸發 】", tasks.Where(t => t.MonthStr == "每月").ToList());
        for (int i = 1; i <= 12; i++) {
            string m = i + "月";
            AddGroup(flow, "【 " + m + " 限定 】", tasks.Where(t => t.MonthStr == m).ToList());
        }
        flow.Controls.Add(new Label() { Height = 100, Text = "" }); // 底部緩衝空間
    }

    private void AddGroup(FlowLayoutPanel container, string header, List<App_RecurringTasks.RecurringTask> subTasks) {
        if (subTasks.Count == 0) return;
        container.Controls.Add(new Label() { Text = header, Font = new Font("Microsoft JhengHei UI", 12f, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(0, 122, 255), Margin = new Padding(0, 20, 0, 10) });
        
        foreach (var t in subTasks) {
            // 【修復】全面改用網格排版 (TableLayoutPanel)，徹底防止按鈕被裁切！
            TableLayoutPanel row = new TableLayoutPanel() { 
                Width = 720, Height = 36, Margin = new Padding(15, 2, 0, 2), 
                ColumnCount = 3, RowCount = 1 
            };
            
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); // [調] 按鈕專屬空間
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); // [✕] 按鈕專屬空間
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // 文字區塊填滿剩餘空間
            
            Button btnEdit = new Button() { Text = "調", Dock = DockStyle.Fill, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 9f), Margin = new Padding(0,0,5,0) };
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (s, e) => { 
                int currentIdx = parentControl.tasks.IndexOf(t);
                if(currentIdx != -1) { new EditRecurringTaskWindow(parentControl, currentIdx, t).ShowDialog(); RefreshData(); }
            };

            Button btnDel = new Button() { Text = "✕", Dock = DockStyle.Fill, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 9f), Margin = new Padding(0,0,5,0) };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += (s, e) => { 
                if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { 
                    parentControl.DeleteTask(t); 
                    RefreshData(); 
                } 
            };

            Label lblItem = new Label() { Text = string.Format("[{0}] {1} {2}", t.TimeStr, t.DateStr, t.Name), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true };
            
            row.Controls.Add(btnEdit, 0, 0); 
            row.Controls.Add(btnDel, 1, 0); 
            row.Controls.Add(lblItem, 2, 0);
            
            container.Controls.Add(row);
        }
    }

    private void ExportToPDF() {
        currentCategoryIndex = 0; currentItemIndex = 0; pageNum = 1;
        PrintDocument pd = new PrintDocument();
        pd.DefaultPageSettings.Margins = new Margins(50, 50, 50, 50);
        pd.PrintPage += Pd_PrintPage;
        PrintDialog prtDlg = new PrintDialog() { Document = pd };
        if (prtDlg.ShowDialog() == DialogResult.OK) pd.Print();
    }

    private void Pd_PrintPage(object sender, PrintPageEventArgs e) {
        Graphics g = e.Graphics;
        float y = e.MarginBounds.Top; float x = e.MarginBounds.Left; float pageBottom = e.MarginBounds.Bottom;
        Font fTitle = new Font("Microsoft JhengHei UI", 18, FontStyle.Bold);
        Font fHeader = new Font("Microsoft JhengHei UI", 13, FontStyle.Bold);
        Font fContent = new Font("Microsoft JhengHei UI", 11);
        Font fFooter = new Font("Microsoft JhengHei UI", 9);

        if (pageNum == 1) {
            g.DrawString("週期任務排程清單 (" + DateTime.Now.ToString("yyyy/MM/dd") + ")", fTitle, Brushes.Black, x, y);
            y += 60;
        }

        var categories = new List<Tuple<string, List<App_RecurringTasks.RecurringTask>>>();
        categories.Add(new Tuple<string, List<App_RecurringTasks.RecurringTask>>("每天觸發", parentControl.tasks.Where(t => t.MonthStr == "每天").ToList()));
        categories.Add(new Tuple<string, List<App_RecurringTasks.RecurringTask>>("每週觸發", parentControl.tasks.Where(t => t.MonthStr == "每週").ToList()));
        categories.Add(new Tuple<string, List<App_RecurringTasks.RecurringTask>>("每月觸發", parentControl.tasks.Where(t => t.MonthStr == "每月").ToList()));
        for (int i = 1; i <= 12; i++) {
            string m = i + "月";
            categories.Add(new Tuple<string, List<App_RecurringTasks.RecurringTask>>(m + " 限定", parentControl.tasks.Where(t => t.MonthStr == m).ToList()));
        }

        for (int i = currentCategoryIndex; i < categories.Count; i++) {
            var cat = categories[i];
            if (cat.Item2.Count == 0) { currentCategoryIndex++; continue; }
            if (currentItemIndex == 0) {
                if (y + 40 > pageBottom) { e.HasMorePages = true; return; }
                g.DrawString("【 " + cat.Item1 + " 】", fHeader, Brushes.Blue, x, y);
                y += 35;
            }
            for (int j = currentItemIndex; j < cat.Item2.Count; j++) {
                if (y + 30 > pageBottom) {
                    currentCategoryIndex = i; currentItemIndex = j;
                    e.HasMorePages = true;
                    g.DrawString("Page " + pageNum, fFooter, Brushes.Gray, e.MarginBounds.Right - 50, pageBottom + 20);
                    pageNum++; return;
                }
                var itm = cat.Item2[j];
                g.DrawString(string.Format("• [{0}] {1} {2}", itm.TimeStr, itm.DateStr, itm.Name), fContent, Brushes.Black, x + 20, y);
                y += 28;
            }
            currentItemIndex = 0; currentCategoryIndex++; y += 15;
        }
        g.DrawString("Page " + pageNum, fFooter, Brushes.Gray, e.MarginBounds.Right - 50, pageBottom + 20);
        e.HasMorePages = false;
    }
}

// ==========================================
// 視窗：新增、調整、設定
// ==========================================
public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parentControl;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate;
    private DateTimePicker dtpTime;

    public AddRecurringTaskWindow(App_RecurringTasks parent) {
        this.parentControl = parent; this.Text = "新增週期任務";
        this.Width = 360; this.Height = 380; this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog;

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(25, 20, 25, 20) };
        flow.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        txtName = new TextBox() { Width = 290, Margin = new Padding(0, 0, 0, 15) }; flow.Controls.Add(txtName);
        flow.Controls.Add(new Label() { Text = "週期類型：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        cmbMonth = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 0, 0, 15) };
        cmbMonth.Items.AddRange(new string[] { "每天", "每週", "每月" });
        for(int i=1;i<=12;i++) cmbMonth.Items.Add(i+"月");
        flow.Controls.Add(cmbMonth);
        flow.Controls.Add(new Label() { Text = "日期條件：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        cmbDate = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 0, 0, 15) };
        flow.Controls.Add(cmbDate);
        cmbMonth.SelectedIndexChanged += (s, e) => {
            cmbDate.Items.Clear();
            if (cmbMonth.Text == "每天") { cmbDate.Items.Add("每日"); cmbDate.Enabled = false; } 
            else if (cmbMonth.Text == "每週") { cmbDate.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); cmbDate.Enabled = true; } 
            else { for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i.ToString()); cmbDate.Items.Add("月底"); cmbDate.Enabled = true; }
            cmbDate.SelectedIndex = 0;
        };
        cmbMonth.SelectedIndex = 0;
        flow.Controls.Add(new Label() { Text = "執行時間：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        
        dtpTime = new DateTimePicker() { Width = 290, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Margin = new Padding(0, 0, 0, 20) };
        dtpTime.Value = DateTime.Today.AddHours(9); 
        
        flow.Controls.Add(dtpTime);
        Button btnSave = new Button() { Text = "建立任務", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnSave.Click += (s, e) => { if (!string.IsNullOrWhiteSpace(txtName.Text)) { parentControl.AddNewTask(txtName.Text, cmbMonth.Text, cmbDate.Text, dtpTime.Value.ToString("HH:mm")); this.Close(); } };
        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}

public class EditRecurringTaskWindow : Form {
    private App_RecurringTasks parentControl;
    private int taskIndex;
    private TextBox txtName;
    private ComboBox cmbMonth, cmbDate;
    private DateTimePicker dtpTime;

    public EditRecurringTaskWindow(App_RecurringTasks parent, int index, App_RecurringTasks.RecurringTask task) {
        this.parentControl = parent; this.taskIndex = index; this.Text = "調整任務條件";
        this.Width = 360; this.Height = 380; this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog;

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(25, 20, 25, 20) };
        flow.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        txtName = new TextBox() { Width = 290, Text = task.Name, Margin = new Padding(0, 0, 0, 15) }; flow.Controls.Add(txtName);
        cmbMonth = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 0, 0, 15) };
        cmbMonth.Items.AddRange(new string[] { "每天", "每週", "每月" });
        for(int i=1;i<=12;i++) cmbMonth.Items.Add(i+"月");
        cmbMonth.Text = task.MonthStr; flow.Controls.Add(cmbMonth);
        cmbDate = new ComboBox() { Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 0, 0, 15) };
        flow.Controls.Add(cmbDate);
        cmbMonth.SelectedIndexChanged += (s, e) => {
            cmbDate.Items.Clear();
            if (cmbMonth.Text == "每天") { cmbDate.Items.Add("每日"); } 
            else if (cmbMonth.Text == "每週") { cmbDate.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); } 
            else { for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i.ToString()); cmbDate.Items.Add("月底"); }
            if (cmbDate.Items.Count > 0) cmbDate.SelectedIndex = 0;
        };
        cmbDate.Text = task.DateStr;
        dtpTime = new DateTimePicker() { Width = 290, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Margin = new Padding(0, 0, 0, 20) };
        DateTime dt; if(DateTime.TryParseExact(task.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dt)) dtpTime.Value = dt;
        flow.Controls.Add(dtpTime);
        Button btnSave = new Button() { Text = "儲存修改", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnSave.Click += (s, e) => { parentControl.UpdateTask(taskIndex, txtName.Text, cmbMonth.Text, cmbDate.Text, dtpTime.Value.ToString("HH:mm")); this.Close(); };
        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}

public class RecurringSettingsWindow : Form {
    private App_RecurringTasks parentControl;
    private ComboBox cmbDigest, cmbAdvanceDays;
    private DateTimePicker dtpDigestTime;

    public RecurringSettingsWindow(App_RecurringTasks parent) {
        this.parentControl = parent; this.Text = "全域排程設定";
        this.Width = 350; this.Height = 230; this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog;

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20) };
        FlowLayoutPanel advRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
        advRow.Controls.Add(new Label() { Text = "所有任務提前", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbAdvanceDays = new ComboBox() { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
        for (int i = 0; i <= 7; i++) cmbAdvanceDays.Items.Add(i.ToString());
        cmbAdvanceDays.Text = parent.advanceDays.ToString();
        advRow.Controls.Add(cmbAdvanceDays);
        advRow.Controls.Add(new Label() { Text = "天加入待辦", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        flow.Controls.Add(advRow);
        FlowLayoutPanel digestRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
        digestRow.Controls.Add(new Label() { Text = "視窗摘要提醒：", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbDigest = new ComboBox() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDigest.Items.AddRange(new string[]{"不提醒", "每週一", "每月1號"}); 
        cmbDigest.Text = parent.digestType;
        dtpDigestTime = new DateTimePicker() { Width = 80, Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true };
        DateTime dt; if(DateTime.TryParseExact(parent.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dt)) dtpDigestTime.Value = dt;
        digestRow.Controls.AddRange(new Control[]{ cmbDigest, dtpDigestTime });
        flow.Controls.Add(digestRow);
        Button btnSave = new Button() { Text = "儲存所有設定", Width = 280, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnSave.Click += (s, e) => { int adv = 0; int.TryParse(cmbAdvanceDays.Text, out adv); parentControl.UpdateGlobalSettings(cmbDigest.Text, dtpDigestTime.Value.ToString("HH:mm"), adv); this.Close(); };
        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}
