/*
 * 檔案功能：週期任務管理與自動推播模組 (支援每日、每週、每月、特定日期及提前推播)
 * 對應選單名稱：週期任務
 * 對應資料庫名稱：(本模組採用純文字檔存儲) MainDB_RecurringTasks.txt
 * 資料表名稱：無 (資料欄位採用 '|' 符號間隔，全域設定採用 #DIGEST 標頭)
 */

using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Threading.Tasks;

public class App_RecurringTasks : UserControl
{
    // --- 核心變數與參考 ---
    private MainForm parentForm;
    private dynamic todoApp; // 接收推播任務的目標清單 (動態綁定 App_TodoList)
    private string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MainDB_RecurringTasks.txt");
    private FlowLayoutPanel taskPanel;
    private System.Windows.Forms.Timer checkTimer;

    // --- 全域排程設定 ---
    public string digestType = "不提醒";
    public string digestTimeStr = "08:00";
    public string lastDigestDate = "";
    public int advanceDays = 0;
    public string scanFrequency = "10分鐘";

    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleOrange = Color.FromArgb(255, 149, 0);
    private static Color AppleGray = Color.FromArgb(142, 142, 147);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);
    private static Font SmallFont = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Regular);

    // --- 資料模型 ---
    public class RecurringTask
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string MonthStr { get; set; } // 每天, 每週, 每月, 特定日期
        public string DateStr { get; set; }  // 1~31, 一~日, yyyy-MM-dd
        public string TimeStr { get; set; }  // HH:mm
        public string LastTriggeredDate { get; set; }
        public string Note { get; set; }
        public string TaskType { get; set; } // 循環, 單次
    }

    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, dynamic todoApp)
    {
        this.parentForm = mainForm;
        this.todoApp = todoApp;

        // 1. 初始化控制項與 DPI 支援
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15);

        // 2. 建構純程式碼 UI
        InitializeUI();

        // 3. 載入資料並啟動計時器 (非同步)
        _ = LoadTasksAndStartTimerAsync();
    }

    /// <summary>
    /// 建構 iOS 風格純程式碼介面 (Code-First UI)
    /// </summary>
    private void InitializeUI()
    {
        // ==========================================
        // 頂部標題與控制區塊
        // ==========================================
        TableLayoutPanel header = new TableLayoutPanel()
        {
            Dock = DockStyle.Top,
            Height = 45,
            ColumnCount = 3,
            BackColor = Color.Transparent
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));

        Label lblTitle = new Label()
        {
            Text = "週期排程任務",
            Font = new Font(MainFont.FontFamily, 14f, FontStyle.Bold),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(5, 0, 0, 0)
        };

        Button btnSet = new Button()
        {
            Text = "全域設定",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = Color.White,
            ForeColor = AppleBlue,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold),
            Margin = new Padding(0, 5, 10, 5) // 內縮並與右側按鈕保持距離
        };
        btnSet.FlatAppearance.BorderSize = 1;
        btnSet.FlatAppearance.BorderColor = AppleBlue;
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).ShowDialog(); };

        Button btnAdd = new Button()
        {
            Text = "新增",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = AppleBlue,
            ForeColor = Color.White,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold),
            Margin = new Padding(0, 5, 0, 5)
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += (s, e) => { new AddEditRecurringTaskWindow(this).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnSet, 1, 0);
        header.Controls.Add(btnAdd, 2, 0);

        // ==========================================
        // 中間列表區塊 (卡片容器)
        // ==========================================
        taskPanel = new FlowLayoutPanel()
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = AppleBgColor
        };

        taskPanel.Resize += (s, e) =>
        {
            int safeWidth = taskPanel.ClientSize.Width - 20;
            if (safeWidth > 0)
            {
                foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
            }
        };

        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent });
        this.Controls.Add(header);
        
        taskPanel.BringToFront();
    }

    /// <summary>
    /// 更新介面顯示 (執行緒安全)
    /// </summary>
    public void RefreshUI()
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(RefreshUI));
            return;
        }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var t in tasks)
        {
            Panel card = new Panel()
            {
                Width = startWidth,
                AutoSize = true,
                MinimumSize = new Size(0, 65),
                Margin = new Padding(0, 0, 0, 15),
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            TableLayoutPanel tlp = new TableLayoutPanel()
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2,
                AutoSize = true
            };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f)); // 備註按鈕
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f)); // 編輯按鈕
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f)); // 刪除按鈕

            // 標題與標籤
            Label lblTitle = new Label()
            {
                Text = $"[{t.TaskType}] {t.Name}",
                Font = BoldFont,
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 5)
            };
            tlp.Controls.Add(lblTitle, 0, 0);

            // 排程時間資訊
            string scheduleStr = t.MonthStr == "特定日期" ? $"{t.DateStr} {t.TimeStr}" : $"{t.MonthStr} {t.DateStr} {t.TimeStr}";
            Label lblDetails = new Label()
            {
                Text = $"觸發條件: {scheduleStr} (上次觸發: {t.LastTriggeredDate})",
                Font = SmallFont,
                ForeColor = Color.DarkGray,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            tlp.Controls.Add(lblDetails, 0, 1);

            // 備註按鈕 (黃色或灰色)
            Button btnNote = new Button()
            {
                Text = "註",
                Dock = DockStyle.Fill,
                BackColor = string.IsNullOrWhiteSpace(t.Note) ? Color.FromArgb(230, 230, 230) : AppleOrange,
                ForeColor = string.IsNullOrWhiteSpace(t.Note) ? Color.Gray : Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = SmallFont,
                Margin = new Padding(5, 5, 5, 5)
            };
            btnNote.FlatAppearance.BorderSize = 0;
            tlp.SetRowSpan(btnNote, 2);
            btnNote.Click += (s, e) => { MessageBox.Show(string.IsNullOrWhiteSpace(t.Note) ? "無備註" : t.Note, $"備註: {t.Name}", MessageBoxButtons.OK, MessageBoxIcon.Information); };
            tlp.Controls.Add(btnNote, 1, 0);

            // 編輯按鈕
            Button btnEdit = new Button()
            {
                Text = "編輯",
                Dock = DockStyle.Fill,
                BackColor = AppleBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = SmallFont,
                Margin = new Padding(5, 5, 5, 5)
            };
            btnEdit.FlatAppearance.BorderSize = 0;
            tlp.SetRowSpan(btnEdit, 2);
            btnEdit.Click += (s, e) => { new AddEditRecurringTaskWindow(this, t).ShowDialog(); };
            tlp.Controls.Add(btnEdit, 2, 0);

            // 刪除按鈕
            Button btnDel = new Button()
            {
                Text = "刪除",
                Dock = DockStyle.Fill,
                BackColor = AppleRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = SmallFont,
                Margin = new Padding(5, 5, 0, 5)
            };
            btnDel.FlatAppearance.BorderSize = 0;
            tlp.SetRowSpan(btnDel, 2);
            btnDel.Click += async (s, e) =>
            {
                if (MessageBox.Show("確定移除此任務？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    tasks.Remove(t);
                    await SaveTasksAsync();
                    RefreshUI();
                }
            };
            tlp.Controls.Add(btnDel, 3, 0);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    // ==========================================
    // 背景排程偵測引擎 (非同步與執行緒安全)
    // ==========================================
    public void UpdateTimerFrequency()
    {
        if (checkTimer != null)
        {
            checkTimer.Interval = GetTimerInterval(scanFrequency);
        }
    }

    private int GetTimerInterval(string freq)
    {
        switch (freq)
        {
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

    private async void CheckTasks()
    {
        DateTime now = DateTime.Now;
        bool needsSave = false;
        List<RecurringTask> toRemove = new List<RecurringTask>();

        foreach (var t in tasks)
        {
            if (TryGetNextTriggerTime(t, now, out DateTime target))
            {
                // 計算是否滿足提前天數
                DateTime triggerThreshold = target.AddDays(-advanceDays);
                if (now >= triggerThreshold)
                {
                    string targetDateStr = target.ToString("yyyy-MM-dd");
                    
                    // 避免重複觸發同一天
                    if (t.LastTriggeredDate != targetDateStr)
                    {
                        string prefix = advanceDays > 0 ? $"[預排-{target:MM/dd}] " : "";
                        
                        // 透過 dynamic 呼叫 App_TodoList 的推播 API (確保 Thread-Safety)
                        if (todoApp != null)
                        {
                            if (parentForm.InvokeRequired)
                            {
                                parentForm.Invoke(new Action(async () => await todoApp.ReceiveTaskAsync(prefix + t.Name, now.ToString("yyyy-MM-dd HH:mm:ss"))));
                            }
                            else
                            {
                                await todoApp.ReceiveTaskAsync(prefix + t.Name, now.ToString("yyyy-MM-dd HH:mm:ss"));
                            }
                        }

                        t.LastTriggeredDate = targetDateStr;
                        needsSave = true;

                        // 若為單次任務則推播後標記刪除
                        if (t.TaskType == "單次") toRemove.Add(t);
                    }
                }
            }
        }

        if (toRemove.Count > 0)
        {
            foreach (var r in toRemove) tasks.Remove(r);
            RefreshUI();
        }

        if (needsSave)
        {
            await SaveTasksAsync();
            RefreshUI();
        }
    }

    /// <summary>
    /// 計算週期任務的下一次觸發時間點
    /// </summary>
    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target)
    {
        target = DateTime.MinValue;
        try
        {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]);
            int m = int.Parse(timeParts[1]);

            if (t.MonthStr == "每天")
            {
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                if (now > target) target = target.AddDays(1);
                return true;
            }
            else if (t.MonthStr == "每週")
            {
                int targetDayOfWeek = "日一二三四五六".IndexOf(t.DateStr);
                int currentDayOfWeek = (int)now.DayOfWeek;
                int daysToAdd = targetDayOfWeek - currentDayOfWeek;
                if (daysToAdd < 0 || (daysToAdd == 0 && now.TimeOfDay.TotalMinutes > h * 60 + m)) daysToAdd += 7;
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0).AddDays(daysToAdd);
                return true;
            }
            else if (t.MonthStr == "每月")
            {
                int validDay = t.DateStr == "月底" ? DateTime.DaysInMonth(now.Year, now.Month) : Math.Min(int.Parse(t.DateStr), DateTime.DaysInMonth(now.Year, now.Month));
                target = new DateTime(now.Year, now.Month, validDay, h, m, 0);
                if (now > target)
                {
                    DateTime nextMonth = now.AddMonths(1);
                    validDay = t.DateStr == "月底" ? DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month) : Math.Min(int.Parse(t.DateStr), DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month));
                    target = new DateTime(nextMonth.Year, nextMonth.Month, validDay, h, m, 0);
                }
                return true;
            }
            else if (t.MonthStr == "特定日期")
            {
                DateTime specDate = DateTime.Parse(t.DateStr);
                target = new DateTime(specDate.Year, specDate.Month, specDate.Day, h, m, 0);
                return true;
            }
        }
        catch { }
        return false;
    }

    // ==========================================
    // 存檔與載入 (非同步 I/O)
    // ==========================================
    private async Task LoadTasksAndStartTimerAsync()
    {
        if (File.Exists(dataFilePath))
        {
            try
            {
                string[] lines = await Task.Run(() => File.ReadAllLines(dataFilePath));
                tasks.Clear();

                foreach (var line in lines)
                {
                    if (line.StartsWith("#DIGEST|"))
                    {
                        var p = line.Split('|');
                        if (p.Length >= 6)
                        {
                            digestType = p[1]; digestTimeStr = p[2]; lastDigestDate = p[3];
                            advanceDays = int.Parse(p[4]); scanFrequency = p[5];
                        }
                        continue;
                    }

                    var parts = line.Split('|');
                    if (parts.Length >= 7)
                    {
                        // 處理 Base64 備註解碼
                        string note = "";
                        try { note = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(parts[5])); } catch { note = parts[5]; }

                        tasks.Add(new RecurringTask
                        {
                            Id = Guid.NewGuid().ToString("N"),
                            Name = parts[0],
                            MonthStr = parts[1],
                            DateStr = parts[2],
                            TimeStr = parts[3],
                            LastTriggeredDate = parts[4],
                            Note = note,
                            TaskType = parts[6]
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入週期任務失敗: {ex.Message}");
            }
        }

        RefreshUI();

        checkTimer = new System.Windows.Forms.Timer();
        UpdateTimerFrequency();
        checkTimer.Tick += (s, e) => CheckTasks();
        checkTimer.Start();
        
        // 啟動時立刻檢查一次
        CheckTasks();
    }

    public async Task SaveTasksAsync()
    {
        try
        {
            List<string> lines = new List<string>();
            lines.Add($"#DIGEST|{digestType}|{digestTimeStr}|{lastDigestDate}|{advanceDays}|{scanFrequency}");
            
            foreach (var t in tasks)
            {
                string base64Note = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(t.Note ?? ""));
                lines.Add($"{t.Name}|{t.MonthStr}|{t.DateStr}|{t.TimeStr}|{t.LastTriggeredDate}|{base64Note}|{t.TaskType}");
            }

            await Task.Run(() => File.WriteAllLines(dataFilePath, lines));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"儲存週期任務失敗: {ex.Message}");
        }
    }
}

// ==========================================
// 視窗：新增/編輯任務 (iOS 風格整合版)
// ==========================================
public class AddEditRecurringTaskWindow : Form
{
    private App_RecurringTasks parent;
    private App_RecurringTasks.RecurringTask editTarget;
    
    private TextBox txtName, txtNote;
    private ComboBox cmbType, cmbCycleType, cmbDate;
    private DateTimePicker dtpSpecificDate, dtpTime;

    public AddEditRecurringTaskWindow(App_RecurringTasks parent, App_RecurringTasks.RecurringTask item = null)
    {
        this.parent = parent;
        this.editTarget = item;

        this.Text = item == null ? "新增週期任務" : "編輯週期任務";
        this.Width = 450;
        this.Height = 580;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel flow = new FlowLayoutPanel()
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20), AutoScroll = true, WrapContents = false
        };

        // --- 輔助方法 ---
        void AddLabel(string text) => flow.Controls.Add(new Label() { Text = text, AutoSize = true, Margin = new Padding(0, 10, 0, 5) });

        // 任務名稱
        AddLabel("任務名稱：");
        txtName = new TextBox() { Width = 380, BorderStyle = BorderStyle.FixedSingle, Text = item?.Name ?? "" };
        flow.Controls.Add(txtName);

        // 執行模式 (循環/單次)
        AddLabel("執行模式：");
        cmbType = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 380 };
        cmbType.Items.AddRange(new string[] { "循環", "單次" });
        cmbType.SelectedItem = item?.TaskType ?? "循環";
        flow.Controls.Add(cmbType);

        // 循環頻率
        AddLabel("循環頻率 / 特定日期：");
        cmbCycleType = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 380 };
        cmbCycleType.Items.AddRange(new string[] { "每天", "每週", "每月", "特定日期" });
        cmbCycleType.SelectedItem = item?.MonthStr ?? "每天";
        flow.Controls.Add(cmbCycleType);

        // 詳細日期 (動態切換)
        AddLabel("細節日期：");
        Panel datePanel = new Panel() { Width = 380, Height = 35, Margin = new Padding(0, 0, 0, 10) };
        cmbDate = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
        dtpSpecificDate = new DateTimePicker() { Format = DateTimePickerFormat.Short, Dock = DockStyle.Fill, Visible = false };
        datePanel.Controls.Add(dtpSpecificDate);
        datePanel.Controls.Add(cmbDate);
        flow.Controls.Add(datePanel);

        // 觸發時間
        AddLabel("觸發時間：");
        dtpTime = new DateTimePicker() { Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Width = 380 };
        if (item != null && DateTime.TryParseExact(item.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtpTime.Value = dtv;
        flow.Controls.Add(dtpTime);

        // 備註
        AddLabel("備註 (選填)：");
        txtNote = new TextBox() { Width = 380, Height = 80, Multiline = true, BorderStyle = BorderStyle.FixedSingle, Text = item?.Note ?? "" };
        flow.Controls.Add(txtNote);

        // 連動邏輯
        cmbCycleType.SelectedIndexChanged += (s, e) =>
        {
            string sel = cmbCycleType.Text;
            if (sel == "特定日期")
            {
                cmbDate.Visible = false;
                dtpSpecificDate.Visible = true;
            }
            else
            {
                dtpSpecificDate.Visible = false;
                cmbDate.Visible = true;
                cmbDate.Items.Clear();
                if (sel == "每天") cmbDate.Items.Add("每日");
                else if (sel == "每週") cmbDate.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" });
                else if (sel == "每月") { for (int i = 1; i <= 31; i++) cmbDate.Items.Add(i.ToString()); cmbDate.Items.Add("月底"); }
                
                if (cmbDate.Items.Count > 0) cmbDate.SelectedIndex = 0;
            }
        };
        // 觸發初始化連動
        cmbCycleType.SelectedIndex = cmbCycleType.Items.IndexOf(item?.MonthStr ?? "每天");
        if (item != null && item.MonthStr != "特定日期" && cmbDate.Items.Contains(item.DateStr)) cmbDate.SelectedItem = item.DateStr;
        if (item != null && item.MonthStr == "特定日期" && DateTime.TryParse(item.DateStr, out DateTime sDate)) dtpSpecificDate.Value = sDate;

        // 儲存按鈕
        Button btnSave = new Button()
        {
            Text = "儲存任務", Width = 380, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), Margin = new Padding(0, 20, 0, 0)
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtName.Text)) { MessageBox.Show("請輸入任務名稱！"); return; }

            string finalDateStr = cmbCycleType.Text == "特定日期" ? dtpSpecificDate.Value.ToString("yyyy-MM-dd") : cmbDate.Text;
            
            if (editTarget == null)
            {
                parent.tasks.Add(new App_RecurringTasks.RecurringTask()
                {
                    Id = Guid.NewGuid().ToString("N"), Name = txtName.Text.Trim(), MonthStr = cmbCycleType.Text,
                    DateStr = finalDateStr, TimeStr = dtpTime.Value.ToString("HH:mm"), LastTriggeredDate = "", Note = txtNote.Text, TaskType = cmbType.Text
                });
            }
            else
            {
                editTarget.Name = txtName.Text.Trim(); editTarget.MonthStr = cmbCycleType.Text;
                editTarget.DateStr = finalDateStr; editTarget.TimeStr = dtpTime.Value.ToString("HH:mm");
                editTarget.Note = txtNote.Text; editTarget.TaskType = cmbType.Text;
            }

            await parent.SaveTasksAsync();
            parent.RefreshUI();
            this.Close();
        };

        flow.Controls.Add(btnSave);
        this.Controls.Add(flow);
    }
}

// ==========================================
// 視窗：全域設定 (iOS 風格)
// ==========================================
public class RecurringSettingsWindow : Form
{
    private App_RecurringTasks parent;
    private NumericUpDown nudAdvance;
    private ComboBox cmbFreq;

    public RecurringSettingsWindow(App_RecurringTasks parent)
    {
        this.parent = parent;
        this.Text = "排程全域設定";
        this.Width = 350; this.Height = 280;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false; this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };

        flow.Controls.Add(new Label() { Text = "提前幾天推播至待辦事項？", AutoSize = true, Margin = new Padding(0, 5, 0, 5) });
        nudAdvance = new NumericUpDown() { Width = 280, Minimum = 0, Maximum = 30, Value = parent.advanceDays };
        flow.Controls.Add(nudAdvance);

        flow.Controls.Add(new Label() { Text = "背景偵測頻率：", AutoSize = true, Margin = new Padding(0, 15, 0, 5) });
        cmbFreq = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 280 };
        cmbFreq.Items.AddRange(new string[] { "即時", "1分鐘", "5分鐘", "10分鐘", "1小時" });
        cmbFreq.SelectedItem = parent.scanFrequency;
        flow.Controls.Add(cmbFreq);

        Button btnSave = new Button()
        {
            Text = "儲存設定", Width = 280, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), Margin = new Padding(0, 25, 0, 0)
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            parent.advanceDays = (int)nudAdvance.Value;
            parent.scanFrequency = cmbFreq.Text;
            parent.UpdateTimerFrequency();
            await parent.SaveTasksAsync();
            this.Close();
        };

        flow.Controls.Add(btnSave);
        this.Controls.Add(flow);
    }
}
