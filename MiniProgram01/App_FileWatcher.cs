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

        InitializeMonitorPanel();
        InitializeSettingsPanel();

        pnlMonitor.Visible = true;
        pnlSettings.Visible = false;

        LoadConfigAndStartWatch();
    }

    private void InitializeMonitorPanel() {
        pnlMonitor = new Panel() { Dock = DockStyle.Fill };
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 3 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85f));

        Label lblTitle = new Label() { Text = "異動紀錄清單", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(10,0,0,0) };
        
        Button btnClear = new Button() { Text = "一鍵清除", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand };
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();

        Button btnGoSet = new Button() { Text = "⚙ 設定", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
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

    private void InitializeSettingsPanel() {
        pnlSettings = new Panel() { Dock = DockStyle.Fill, BackColor = Color.White };
        FlowLayoutPanel setFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20), AutoScroll = true };
        
        Button btnBack = new Button() { Text = "⬅ 返回紀錄列表", Width = 130, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.Gray, ForeColor = Color.White, Cursor = Cursors.Hand };
        btnBack.Click += (s, e) => { pnlMonitor.Visible = true; pnlSettings.Visible = false; };
        setFlow.Controls.Add(btnBack);

        setFlow.Controls.Add(new Label() { Text = "【新路徑設定】", Font = new Font(MainFont.FontFamily, 11f, FontStyle.Bold), Margin = new Padding(0, 20, 0, 10) });

        txtSource = AddPathRowWithPaste(setFlow, "來源路徑：");
        txtBackup = AddPathRowWithPaste(setFlow, "備份路徑：");

        cmbMethod = AddComboRow(setFlow, "監控方式：", new string[] { "顯示在監控", "背景執行" }, false);
        cmbFreq = AddComboRow(setFlow, "掃描頻率：", new string[] { "1秒", "3秒", "5秒", "10秒" }, true);
        cmbDepth = AddComboRow(setFlow, "監控深度：", new string[] { "無限層", "僅本層", "第一層", "第二層", "第三層" }, true);

        Button btnAdd = new Button() { Text = "+ 新增至監控任務", Width = 320, Height = 40, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(0,15,0,15), Cursor = Cursors.Hand };
        btnAdd.Click += (s, e) => AddNewTask();
        setFlow.Controls.Add(btnAdd);

        setFlow.Controls.Add(new Label() { Text = "【管理中心】", Font = new Font(MainFont.FontFamily, 11f, FontStyle.Bold), Margin = new Padding(0, 20, 0, 10) });
        Button btnShowList = new Button() { Text = "📋 開啟監控項目清冊", Width = 320, Height = 45, FlatStyle = FlatStyle.Flat, BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand, Font = new Font(MainFont, FontStyle.Bold) };
        btnShowList.Click += (s, e) => { 
            using(var listForm = new MonitorListWindow(pathPairs, (key) => { pathPairs.Remove(key); SaveAllConfigs(); })) { 
                listForm.ShowDialog(); 
            } 
        };
        setFlow.Controls.Add(btnShowList);

        pnlSettings.Controls.Add(setFlow);
        this.Controls.Add(pnlSettings);
    }

    private TextBox AddPathRowWithPaste(FlowLayoutPanel container, string labelText) {
        container.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        TextBox tb = new TextBox() { Width = 210, Font = MainFont };
        Button btnSel = new Button() { Text = "選取", Width = 45, Height = 25 };
        btnSel.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; };
        Button btnPaste = new Button() { Text = "貼上", Width = 45, Height = 25 };
        btnPaste.Click += (s, e) => { string p = Clipboard.GetText().Trim(' ', '\"'); if (!string.IsNullOrEmpty(p)) tb.Text = p; };
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

    private void AddNewTask() {
        string src = txtSource.Text.Trim();
        if (!Directory.Exists(src)) { MessageBox.Show("來源路徑無效！"); return; }
        pathPairs[src] = string.Format("{0}|{1}|{2}|{3}|{4}", src, txtBackup.Text.Trim(), cmbMethod.Text, cmbFreq.Text, cmbDepth.Text);
        SaveAllConfigs(); StartWatcherFromLine(pathPairs[src]);
        txtSource.Text = ""; txtBackup.Text = "";
        MessageBox.Show("監控任務已建立！");
    }

    private void StartWatcherFromLine(string line) {
        var p = line.Split('|'); if (p.Length < 5) return;
        FileSystemWatcher w = new FileSystemWatcher(p[0]) { IncludeSubdirectories = (p[4] != "僅本層"), EnableRaisingEvents = true };
        w.Changed += OnFileEvent; w.Created += OnFileEvent; watchers.Add(w);
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        string src = ((FileSystemWatcher)sender).Path;
        if (!pathPairs.ContainsKey(src)) return;
        var p = pathPairs[src].Split('|');
        int targetDepth = ParseDepth(p[4]);
        if (targetDepth != -1 && GetPathDepth(src, e.FullPath) > targetDepth) return;
        lock (lockObj) {
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (DateTime.Now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = DateTime.Now;
        }
        DoBackup(src, e.FullPath, p[1], p[2] == "顯示在監控", p[3]);
    }

    private int GetPathDepth(string root, string file) {
        string rel = file.Replace(root, "").TrimStart('\\', '/');
        return string.IsNullOrEmpty(rel) ? 0 : rel.Split(new char[] { '\\', '/' }).Length - 1;
    }

    private int ParseDepth(string d) {
        if (d == "無限層") return -1; if (d == "僅本層") return 0;
        if (d == "第一層") return 1; if (d == "第二層") return 2; if (d == "第三層") return 3;
        int res; return int.TryParse(d.Replace("層", ""), out res) ? res : -1;
    }

    private void DoBackup(string src, string file, string dst, bool showUI, string freqStr) {
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

// ==========================================
// 【優化後的清冊視窗】大尺寸、居中、無遮擋
// ==========================================
public class MonitorListWindow : Form {
    public MonitorListWindow(Dictionary<string, string> data, Action<string> onDelete) {
        this.Text = "📋 監控項目管理中心 (顯示所有設定)";
        this.Width = 800; // 加寬
        this.Height = 850; // 加高，確保能容納 10 個不滾動
        this.StartPosition = FormStartPosition.CenterScreen; // 螢幕正中央
        this.BackColor = Color.White;
        this.Font = new Font("Microsoft JhengHei UI", 10f);

        FlowLayoutPanel list = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15), 
            FlowDirection = FlowDirection.TopDown, WrapContents = false 
        };
        
        foreach (var key in data.Keys) {
            var p = data[key].Split('|');
            // 卡片高度稍微壓縮，確保能一屏顯示更多
            Panel card = new Panel() { Width = 740, Height = 82, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 10), BackColor = Color.FromArgb(248, 248, 250) };
            
            Label lblSrc = new Label() { Text = "源：" + p[0], Location = new Point(10, 8), Width = 640, ForeColor = Color.FromArgb(0, 102, 204), AutoSize = false, Height = 22 };
            Label lblDst = new Label() { Text = "備：" + p[1], Location = new Point(10, 30), Width = 640, ForeColor = Color.FromArgb(0, 153, 76), AutoSize = false, Height = 22 };
            Label lblSet = new Label() { Text = string.Format("模式：{0} | 頻率：{1} | 深度：{2}", p[2], p[3], p[4]), Location = new Point(10, 54), Width = 500, Font = new Font("Microsoft JhengHei UI", 8.5f), ForeColor = Color.DimGray };
            
            Button btnDel = new Button() { 
                Text = "移除任務", Location = new Point(655, 23), Width = 75, Height = 35, 
                BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand 
            };
            btnDel.Click += (s, e) => { if(MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { onDelete(key); this.Close(); } };
            
            card.Controls.AddRange(new Control[] { lblSrc, lblDst, lblSet, btnDel });
            list.Controls.Add(card);
        }
        this.Controls.Add(list);
    }
}
