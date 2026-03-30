using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private Panel pnlMonitor, pnlSettings;
    private FlowLayoutPanel cardPanel, configListPanel;
    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private Dictionary<string, string> pathPairs = new Dictionary<string, string>();
    private Dictionary<string, DateTime> lastProcessedTimes = new Dictionary<string, DateTime>();
    private readonly object lockObj = new object();
    private TextBox txtSource, txtBackup;
    private ComboBox cmbMethod, cmbFreq, cmbDepth;
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9f);

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        // 初始化兩個面板
        pnlMonitor = new Panel() { Dock = DockStyle.Fill };
        pnlSettings = new Panel() { Dock = DockStyle.Fill, Visible = false, BackColor = Color.White };

        // --- 監控主介面 ---
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 3 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // 標題占滿
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85f));  // 清除按鈕固定寬
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75f));  // 設定按鈕固定寬

        Label lblTitle = new Label() { Text = "異動紀錄", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5,0,0,0) };
        Button btnClear = new Button() { Text = "一鍵清除", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8) };
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();
        Button btnGoSet = new Button() { Text = "⚙ 設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,5,8) };
        btnGoSet.Click += (s, e) => { pnlMonitor.Visible = false; pnlSettings.Visible = true; };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnClear, 1, 0);
        header.Controls.Add(btnGoSet, 2, 0);
        pnlMonitor.Controls.Add(header);

        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        pnlMonitor.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        // --- 設定介面 (同前一版本邏輯，含資料夾選取與清冊) ---
        InitializeSettingsUI();

        this.Controls.Add(pnlMonitor);
        this.Controls.Add(pnlSettings);
        LoadConfigAndStartWatch();
    }

    private void InitializeSettingsUI() {
        FlowLayoutPanel setFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10), AutoScroll = true };
        Button btnBack = new Button() { Text = "⬅ 返回紀錄", Width = 100, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = Color.Gray, ForeColor = Color.White };
        btnBack.Click += (s, e) => { pnlMonitor.Visible = true; pnlSettings.Visible = false; };
        setFlow.Controls.Add(btnBack);
        setFlow.Controls.Add(new Label() { Text = "【監控設定】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) });

        txtSource = AddPathRow(setFlow, "來源路徑：");
        txtBackup = AddPathRow(setFlow, "備份路徑：");
        cmbMethod = AddComboRow(setFlow, "監控方式：", new string[] { "顯示在監控", "背景執行" }, false);
        cmbFreq = AddComboRow(setFlow, "掃描頻率：", new string[] { "1秒", "3秒", "5秒", "10秒" }, true);
        cmbDepth = AddComboRow(setFlow, "監控深度：", new string[] { "僅本層", "無限層" }, true);

        Button btnAdd = new Button() { Text = "+ 新增監控項目", Width = 310, Height = 35, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold) };
        btnAdd.Click += (s, e) => AddNewTask();
        setFlow.Controls.Add(btnAdd);
        
        setFlow.Controls.Add(new Label() { Text = "【監控清冊】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 10, 0, 5) });
        configListPanel = new FlowLayoutPanel() { Width = 320, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        setFlow.Controls.Add(configListPanel);
        pnlSettings.Controls.Add(setFlow);
    }

    private TextBox AddPathRow(FlowLayoutPanel container, string labelText) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        TextBox tb = new TextBox() { Width = 170 };
        Button btn = new Button() { Text = "...", Width = 30, Height = 25 };
        btn.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; };
        row.Controls.AddRange(new Control[] { tb, btn });
        container.Controls.Add(row);
        return tb;
    }

    private ComboBox AddComboRow(FlowLayoutPanel container, string labelText, string[] items, bool editable) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        ComboBox cb = new ComboBox() { Width = 140, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        cb.Items.AddRange(items); cb.SelectedIndex = 0;
        row.Controls.Add(cb);
        container.Controls.Add(row);
        return cb;
    }

    private void AddNewTask() {
        string src = txtSource.Text.Trim();
        if (!Directory.Exists(src)) return;
        string line = string.Format("{0}|{1}|{2}|{3}|{4}", src, txtBackup.Text.Trim(), cmbMethod.Text, cmbFreq.Text, cmbDepth.Text);
        pathPairs[src] = line; SaveAllConfigs(); StartWatcherFromLine(line); RefreshConfigListUI();
        txtSource.Text = ""; txtBackup.Text = "";
    }

    private void StartWatcherFromLine(string line) {
        var p = line.Split('|'); if (p.Length < 5) return;
        FileSystemWatcher w = new FileSystemWatcher(p[0]) { IncludeSubdirectories = (p[4] == "無限層"), EnableRaisingEvents = true };
        w.Changed += OnFileEvent; w.Created += OnFileEvent; watchers.Add(w);
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        lock (lockObj) {
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (DateTime.Now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = DateTime.Now;
        }
        string src = ((FileSystemWatcher)sender).Path;
        if (pathPairs.ContainsKey(src)) {
            var p = pathPairs[src].Split('|');
            DoBackup(src, e.FullPath, p[1], p[2] == "顯示在監控");
        }
    }

    private void DoBackup(string src, string file, string dst, bool showUI) {
        try {
            Uri rootUri = new Uri(src.EndsWith("\\") ? src : src + "\\");
            string rel = Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(file)).ToString().Replace('/', '\\'));
            string target = Path.Combine(dst, rel);
            if (!Directory.Exists(Path.GetDirectoryName(target))) Directory.CreateDirectory(Path.GetDirectoryName(target));
            System.Threading.Thread.Sleep(500); File.Copy(file, target, true);
            if (showUI) {
                this.Invoke(new Action(() => {
                    Panel c = new Panel() { Width = 330, Height = 65, BackColor = Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 2, 0, 2) };
                    c.Controls.Add(new Label() { Text = Path.GetFileName(file) + "\n路徑: " + rel, Location = new Point(5, 5), Width = 240, Height = 40 });
                    Button b = new Button() { Text = "查看", Location = new Point(255, 12), Width = 55, Height = 30, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                    b.Click += (s, e2) => System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + file + "\"");
                    c.Controls.Add(b); cardPanel.Controls.Add(c); cardPanel.Controls.SetChildIndex(c, 0);
                    if (cardPanel.Controls.Count > 15) cardPanel.Controls.RemoveAt(15);
                }));
            }
        } catch { }
    }

    private void RefreshConfigListUI() {
        configListPanel.Controls.Clear();
        foreach (var key in pathPairs.Keys) {
            var p = pathPairs[key].Split('|');
            Panel row = new Panel() { Width = 310, Height = 50, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 2, 0, 2) };
            row.Controls.Add(new Label() { Text = "源:" + Path.GetFileName(p[0]) + "\n備:" + Path.GetFileName(p[1]), Location = new Point(5, 5), Width = 230, Font = new Font(MainFont.FontFamily, 8f) });
            Button btnDel = new Button() { Text = "刪除", Location = new Point(245, 10), Width = 55, Height = 28, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnDel.Click += (s, e) => { pathPairs.Remove(key); SaveAllConfigs(); RefreshConfigListUI(); };
            row.Controls.Add(btnDel); configListPanel.Controls.Add(row);
        }
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            var parts = line.Split('|'); if (parts.Length >= 5 && Directory.Exists(parts[0])) { pathPairs[parts[0]] = line; StartWatcherFromLine(line); }
        }
        RefreshConfigListUI();
    }

    private void SaveAllConfigs() {
        List<string> lines = new List<string>(); foreach (var val in pathPairs.Values) lines.Add(val);
        File.WriteAllLines(configFile, lines);
    }
}
