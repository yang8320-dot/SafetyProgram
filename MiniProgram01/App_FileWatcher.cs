using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    
    // UI 面板
    private Panel pnlMonitor;  // 顯示監控紀錄的面板
    private Panel pnlSettings; // 隱藏的設定面板
    private FlowLayoutPanel cardPanel;
    private FlowLayoutPanel configListPanel; // 設定頁面中的清冊

    // 監控核心
    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private Dictionary<string, string> pathPairs = new Dictionary<string, string>();
    private Dictionary<string, DateTime> lastProcessedTimes = new Dictionary<string, DateTime>();
    private readonly object lockObj = new object();

    // 設定元件
    private TextBox txtSource, txtBackup;
    private ComboBox cmbMethod, cmbFreq, cmbDepth;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9f);

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        InitializeMonitorPanel();
        InitializeSettingsPanel();

        // 預設顯示監控面板
        pnlMonitor.Visible = true;
        pnlSettings.Visible = false;

        LoadConfigAndStartWatch();
    }

    // --- [1] 監控面板介面 (主畫面) ---
    private void InitializeMonitorPanel() {
        pnlMonitor = new Panel() { Dock = DockStyle.Fill };
        
        // 頂部控制列
        FlowLayoutPanel header = new FlowLayoutPanel() { Dock = DockStyle.Top, Height = 40, WrapContents = false };
        Label lblTitle = new Label() { Text = "🔍 異動紀錄", Font = new Font(MainFont, FontStyle.Bold), AutoSize = true, Margin = new Padding(5, 10, 50, 0) };
        
        Button btnClear = new Button() { Text = "🗑️ 清除清單", Width = 80, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.White };
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();

        Button btnGoSet = new Button() { Text = "⚙️ 設定", Width = 70, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray };
        btnGoSet.Click += (s, e) => { pnlMonitor.Visible = false; pnlSettings.Visible = true; };

        header.Controls.AddRange(new Control[] { lblTitle, btnClear, btnGoSet });
        pnlMonitor.Controls.Add(header);

        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        pnlMonitor.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        this.Controls.Add(pnlMonitor);
    }

    // --- [2] 設定面板介面 (隱藏分頁) ---
    private void InitializeSettingsPanel() {
        pnlSettings = new Panel() { Dock = DockStyle.Fill, BackColor = Color.White };
        
        FlowLayoutPanel setFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10), AutoScroll = true };
        
        // 返回按鈕
        Button btnBack = new Button() { Text = "⬅ 返回監控列表", Width = 120, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = Color.DimGray, ForeColor = Color.White };
        btnBack.Click += (s, e) => { pnlMonitor.Visible = true; pnlSettings.Visible = false; };
        setFlow.Controls.Add(btnBack);

        setFlow.Controls.Add(new Label() { Text = "【監控設定】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) });

        // 來源與備份路徑 (帶有點選按鈕)
        txtSource = AddPathRow(setFlow, "來源路徑：");
        txtBackup = AddPathRow(setFlow, "備份路徑：");

        // 監控方式、頻率、深度
        cmbMethod = AddComboRow(setFlow, "監控方式：", new string[] { "顯示在監控", "背景執行" }, false);
        cmbFreq = AddComboRow(setFlow, "掃描頻率：", new string[] { "1秒", "3秒", "5秒", "10秒" }, true);
        cmbDepth = AddComboRow(setFlow, "監控深度：", new string[] { "僅本層", "無限層" }, true);

        Button btnAddConfig = new Button() { Text = "+ 新增至監控清冊", Width = 310, Height = 35, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0,10,0,10) };
        btnAddConfig.Click += (s, e) => AddNewMonitorTask();
        setFlow.Controls.Add(btnAddConfig);

        setFlow.Controls.Add(new Label() { Text = "【監控項目清冊】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 10, 0, 5) });
        
        configListPanel = new FlowLayoutPanel() { Width = 320, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        setFlow.Controls.Add(configListPanel);

        pnlSettings.Controls.Add(setFlow);
        this.Controls.Add(pnlSettings);
    }

    // 輔助 UI: 路徑輸入行
    private TextBox AddPathRow(FlowLayoutPanel container, string labelText) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 3, 0, 3) };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        TextBox tb = new TextBox() { Width = 180 };
        Button btn = new Button() { Text = "...", Width = 35, Height = 25 };
        btn.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; };
        row.Controls.AddRange(new Control[] { tb, btn });
        container.Controls.Add(row);
        return tb;
    }

    // 輔助 UI: 下拉選單行
    private ComboBox AddComboRow(FlowLayoutPanel container, string labelText, string[] items, bool editable) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 3, 0, 3) };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        ComboBox cb = new ComboBox() { Width = 150, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        cb.Items.AddRange(items); if(items.Length > 0) cb.SelectedIndex = 0;
        row.Controls.Add(cb);
        container.Controls.Add(row);
        return cb;
    }

    // --- [3] 核心邏輯處理 ---

    private void AddNewMonitorTask() {
        string src = txtSource.Text.Trim();
        string dst = txtBackup.Text.Trim();
        if (!Directory.Exists(src)) { MessageBox.Show("來源路徑無效"); return; }
        
        string configLine = string.Format("{0}|{1}|{2}|{3}|{4}", src, dst, cmbMethod.Text, cmbFreq.Text, cmbDepth.Text);
        pathPairs[src] = configLine;
        SaveAllConfigs();
        StartWatcherFromLine(configLine);
        RefreshConfigListUI();
        txtSource.Text = ""; txtBackup.Text = "";
        MessageBox.Show("已成功加入監控清冊");
    }

    private void StartWatcherFromLine(string line) {
        var p = line.Split('|');
        if (p.Length < 5) return;
        FileSystemWatcher w = new FileSystemWatcher(p[0]) { 
            IncludeSubdirectories = (p[4] == "無限層"), 
            EnableRaisingEvents = true,
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName 
        };
        w.Changed += OnFileEvent; w.Created += OnFileEvent;
        watchers.Add(w);
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        lock (lockObj) {
            DateTime now = DateTime.Now;
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = now;
        }
        string src = ((FileSystemWatcher)sender).Path;
        if (pathPairs.ContainsKey(src)) {
            var p = pathPairs[src].Split('|');
            // 判斷監控方式
            bool showUI = p[2] == "顯示在監控";
            DoBackup(src, e.FullPath, p[1], showUI);
        }
    }

    private void DoBackup(string src, string file, string dst, bool showUI) {
        try {
            Uri rootUri = new Uri(src.EndsWith("\\") ? src : src + "\\");
            string rel = Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(file)).ToString().Replace('/', '\\'));
            string target = Path.Combine(dst, rel);
            string targetDir = Path.GetDirectoryName(target);
            if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
            
            System.Threading.Thread.Sleep(500);
            File.Copy(file, target, true);

            if (showUI) {
                this.Invoke(new Action(() => {
                    Panel c = new Panel() { Width = 330, Height = 70, BackColor = Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 3, 0, 3) };
                    c.Controls.Add(new Label() { Text = Path.GetFileName(file) + "\n" + rel, Location = new Point(5, 5), Width = 230, Height = 40 });
                    Button b = new Button() { Text = "查看", Location = new Point(255, 15), Width = 55, Height = 30, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                    b.Click += (s, e2) => System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + file + "\"");
                    c.Controls.Add(b);
                    cardPanel.Controls.Add(c); cardPanel.Controls.SetChildIndex(c, 0);
                    if (cardPanel.Controls.Count > 20) cardPanel.Controls.RemoveAt(20);
                }));
            }
        } catch { }
    }

    private void RefreshConfigListUI() {
        configListPanel.Controls.Clear();
        foreach (var key in pathPairs.Keys) {
            var p = pathPairs[key].Split('|');
            Panel row = new Panel() { Width = 310, Height = 55, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 2, 0, 2) };
            Label lbl = new Label() { Text = string.Format("源:{0}\n備:{1}", Path.GetFileName(p[0]), Path.GetFileName(p[1])), Location = new Point(5, 5), Width = 230, Height = 45, Font = new Font(MainFont.FontFamily, 8f) };
            Button btnDel = new Button() { Text = "刪除", Location = new Point(245, 12), Width = 55, Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnDel.Click += (s, e) => { pathPairs.Remove(key); SaveAllConfigs(); RefreshConfigListUI(); MessageBox.Show("已移除監控項目，重啟後生效"); };
            row.Controls.Add(lbl); row.Controls.Add(btnDel);
            configListPanel.Controls.Add(row);
        }
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;
            var parts = line.Split('|');
            if (parts.Length >= 5 && Directory.Exists(parts[0])) {
                pathPairs[parts[0]] = line;
                StartWatcherFromLine(line);
            }
        }
        RefreshConfigListUI();
    }

    private void SaveAllConfigs() {
        List<string> lines = new List<string>() { "# 來源|備份|方式|頻率|深度" };
        foreach (var val in pathPairs.Values) lines.Add(val);
        File.WriteAllLines(configFile, lines);
    }
}
