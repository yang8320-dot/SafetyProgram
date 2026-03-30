using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private Panel pnlMonitor, pnlSettings;
    private FlowLayoutPanel cardPanel;

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

        pnlMonitor.Visible = true;
        pnlSettings.Visible = false;

        LoadConfigAndStartWatch();
    }

    // --- [1] 監控主面板 (顯示紀錄) ---
    private void InitializeMonitorPanel() {
        pnlMonitor = new Panel() { Dock = DockStyle.Fill };
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 3 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));

        Label lblTitle = new Label() { Text = "異動紀錄清單", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5,0,0,0) };
        
        Button btnClear = new Button() { Text = "一鍵清除", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand };
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();

        Button btnGoSet = new Button() { Text = "⚙ 設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,5,8), Cursor = Cursors.Hand };
        btnGoSet.Click += (s, e) => { pnlMonitor.Visible = false; pnlSettings.Visible = true; };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnClear, 1, 0);
        header.Controls.Add(btnGoSet, 2, 0);
        pnlMonitor.Controls.Add(header);

        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        pnlMonitor.Controls.Add(cardPanel);
        cardPanel.BringToFront();
        this.Controls.Add(pnlMonitor);
    }

    // --- [2] 設定面板 (參數與路徑) ---
    private void InitializeSettingsPanel() {
        pnlSettings = new Panel() { Dock = DockStyle.Fill, BackColor = Color.White };
        FlowLayoutPanel setFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(15), AutoScroll = true };
        
        Button btnBack = new Button() { Text = "⬅ 返回紀錄列表", Width = 120, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = Color.Gray, ForeColor = Color.White };
        btnBack.Click += (s, e) => { pnlMonitor.Visible = true; pnlSettings.Visible = false; };
        setFlow.Controls.Add(btnBack);

        setFlow.Controls.Add(new Label() { Text = "【新路徑設定】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) });

        txtSource = AddPathRowWithPaste(setFlow, "來源路徑：");
        txtBackup = AddPathRowWithPaste(setFlow, "備份路徑：");

        cmbMethod = AddComboRow(setFlow, "監控方式：", new string[] { "顯示在監控", "背景執行" }, false);
        cmbFreq = AddComboRow(setFlow, "掃描頻率：", new string[] { "1秒", "3秒", "5秒", "10秒", "自定秒數" }, true);
        cmbDepth = AddComboRow(setFlow, "監控深度：", new string[] { "無限層", "僅本層", "第一層", "第二層", "第三層", "自定層數" }, true);

        Button btnAdd = new Button() { Text = "+ 新增至監控任務", Width = 315, Height = 38, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0,10,0,10) };
        btnAdd.Click += (s, e) => AddNewTask();
        setFlow.Controls.Add(btnAdd);

        setFlow.Controls.Add(new Label() { Text = "【管理中心】", Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0, 15, 0, 5) });
        Button btnShowList = new Button() { Text = "📋 開啟監控項目清冊", Width = 315, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.WhiteSmoke };
        btnShowList.Click += (s, e) => { using(var listForm = new MonitorListWindow(pathPairs, (key) => { pathPairs.Remove(key); SaveAllConfigs(); })) { listForm.ShowDialog(); } };
        setFlow.Controls.Add(btnShowList);

        pnlSettings.Controls.Add(setFlow);
        this.Controls.Add(pnlSettings);
    }

    // 輔助 UI: 帶有「選取」與「貼上」的路徑行
    private TextBox AddPathRowWithPaste(FlowLayoutPanel container, string labelText) {
        container.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        TextBox tb = new TextBox() { Width = 215, Font = MainFont };
        Button btnSel = new Button() { Text = "選取", Width = 45, Height = 25 };
        btnSel.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; };
        Button btnPaste = new Button() { Text = "貼上", Width = 45, Height = 25 };
        btnPaste.Click += (s, e) => { 
            string p = Clipboard.GetText().Trim(' ', '\"'); // 自動修剪空格與引號
            if (!string.IsNullOrEmpty(p)) tb.Text = p;
        };
        row.Controls.AddRange(new Control[] { tb, btnSel, btnPaste });
        container.Controls.Add(row);
        return tb;
    }

    private ComboBox AddComboRow(FlowLayoutPanel container, string labelText, string[] items, bool editable) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 5, 0, 5) };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        ComboBox cb = new ComboBox() { Width = 150, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        cb.Items.AddRange(items); cb.SelectedIndex = 0;
        row.Controls.Add(cb);
        container.Controls.Add(row);
        return cb;
    }

    // --- [3] 核心引擎 ---

    private void AddNewTask() {
        string src = txtSource.Text.Trim();
        if (!Directory.Exists(src)) { MessageBox.Show("來源路徑無效！"); return; }
        string line = string.Format("{0}|{1}|{2}|{3}|{4}", src, txtBackup.Text.Trim(), cmbMethod.Text, cmbFreq.Text, cmbDepth.Text);
        pathPairs[src] = line; SaveAllConfigs(); StartWatcherFromLine(line);
        txtSource.Text = ""; txtBackup.Text = "";
        MessageBox.Show("監控任務已建立！");
    }

    private void StartWatcherFromLine(string line) {
        var p = line.Split('|'); if (p.Length < 5) return;
        FileSystemWatcher w = new FileSystemWatcher(p[0]) { 
            IncludeSubdirectories = (p[4] != "僅本層"), // 只要不是僅本層，底層就開啓遞迴，由事件端過濾深度
            EnableRaisingEvents = true 
        };
        w.Changed += OnFileEvent; w.Created += OnFileEvent; watchers.Add(w);
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        string src = ((FileSystemWatcher)sender).Path;
        if (!pathPairs.ContainsKey(src)) return;
        var p = pathPairs[src].Split('|');

        // 【深度過濾邏輯】
        int targetDepth = ParseDepth(p[4]);
        if (targetDepth != -1) { // -1 代表無限層
            int currentDepth = GetPathDepth(src, e.FullPath);
            if (currentDepth > targetDepth) return;
        }

        lock (lockObj) {
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (DateTime.Now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = DateTime.Now;
        }
        DoBackup(src, e.FullPath, p[1], p[2] == "顯示在監控", p[3]);
    }

    private int GetPathDepth(string root, string file) {
        string rel = file.Replace(root, "").TrimStart('\\', '/');
        if (string.IsNullOrEmpty(rel)) return 0;
        return rel.Split(new char[] { '\\', '/' }).Length - 1;
    }

    private int ParseDepth(string d) {
        if (d == "無限層") return -1;
        if (d == "僅本層") return 0;
        if (d == "第一層") return 1;
        if (d == "第二層") return 2;
        if (d == "第三層") return 3;
        int res; if (int.TryParse(d.Replace("層", ""), out res)) return res;
        return -1;
    }

    private void DoBackup(string src, string file, string dst, bool showUI, string freqStr) {
        try {
            Uri rootUri = new Uri(src.EndsWith("\\") ? src : src + "\\");
            string rel = Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(file)).ToString().Replace('/', '\\'));
            string target = Path.Combine(dst, rel);
            if (!Directory.Exists(Path.GetDirectoryName(target))) Directory.CreateDirectory(Path.GetDirectoryName(target));
            
            int delay = 500;
            int customSec; if (int.TryParse(freqStr.Replace("秒", ""), out customSec)) delay = Math.Max(500, customSec * 100);

            System.Threading.Thread.Sleep(delay);
            File.Copy(file, target, true);

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

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            var parts = line.Split('|'); if (parts.Length >= 5 && Directory.Exists(parts[0])) { pathPairs[parts[0]] = line; StartWatcherFromLine(line); }
        }
    }

    private void SaveAllConfigs() {
        List<string> lines = new List<string>(); foreach (var val in pathPairs.Values) lines.Add(val);
        File.WriteAllLines(configFile, lines);
    }
}

// --- [4] 獨立視窗：監控項目清冊 ---
public class MonitorListWindow : Form {
    public MonitorListWindow(Dictionary<string, string> data, Action<string> onDelete) {
        this.Text = "📋 監控項目清冊管理";
        this.Width = 450; this.Height = 500;
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;

        FlowLayoutPanel list = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        
        foreach (var key in data.Keys) {
            var p = data[key].Split('|');
            Panel card = new Panel() { Width = 400, Height = 110, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 5, 0, 5), BackColor = Color.FromArgb(245, 245, 247) };
            
            Label lblSrc = new Label() { Text = "源：" + p[0], Location = new Point(10, 10), Width = 380, AutoEllipsis = false, Height = 35, ForeColor = Color.Navy };
            Label lblDst = new Label() { Text = "備：" + p[1], Location = new Point(10, 45), Width = 380, Height = 35, ForeColor = Color.DarkGreen };
            Label lblSet = new Label() { Text = string.Format("模式：{0} | 頻率：{1} | 深度：{2}", p[2], p[3], p[4]), Location = new Point(10, 82), Width = 280, Font = new Font("Microsoft JhengHei UI", 8f) };
            
            Button btnDel = new Button() { Text = "刪除任務", Location = new Point(310, 78), Width = 75, Height = 25, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnDel.Click += (s, e) => { if(MessageBox.Show("確定移除此監控？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { onDelete(key); this.Close(); } };
            
            card.Controls.AddRange(new Control[] { lblSrc, lblDst, lblSet, btnDel });
            list.Controls.Add(card);
        }
        this.Controls.Add(list);
    }
}
