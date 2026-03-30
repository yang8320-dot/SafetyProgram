using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel cardPanel;
    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private Dictionary<string, string> pathPairs = new Dictionary<string, string>();
    private Dictionary<string, DateTime> lastProcessedTimes = new Dictionary<string, DateTime>();
    private readonly object lockObj = new object();

    // UI 元件
    private TextBox txtSource, txtBackup;
    private ComboBox cmbFreq, cmbDepth;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9f);

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(8);

        // --- 總容器 (由上往下自動排列) ---
        FlowLayoutPanel mainPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false 
        };

        // 1. 新增路徑區
        FlowLayoutPanel r1 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0,0,0,5) };
        txtSource = new TextBox() { Width = 135, Font = MainFont, PlaceholderText = "來源路徑" };
        txtBackup = new TextBox() { Width = 135, Font = MainFont, PlaceholderText = "備份路徑" };
        Button btnAdd = new Button() { Text = "+", Width = 35, Height = 25, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        btnAdd.Click += (s, e) => AddNewPath(txtSource.Text, txtBackup.Text);
        r1.Controls.AddRange(new Control[] { txtSource, txtBackup, btnAdd });

        // 2. 參數設定區 (頻率與深度)
        FlowLayoutPanel r2 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0,0,0,5) };
        r2.Controls.Add(new Label() { Text = "頻率:", AutoSize = true, Margin = new Padding(0, 5, 0, 0) });
        cmbFreq = new ComboBox() { Width = 65, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbFreq.Items.AddRange(new string[] { "1秒", "3秒", "5秒", "10秒" }); cmbFreq.SelectedIndex = 1;
        
        r2.Controls.Add(new Label() { Text = "深度:", AutoSize = true, Margin = new Padding(5, 5, 0, 0) });
        cmbDepth = new ComboBox() { Width = 75, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbDepth.Items.AddRange(new string[] { "僅本層", "無限層" }); cmbDepth.SelectedIndex = 1;

        Button btnSave = new Button() { Text = "套用設定", Width = 70, Height = 25, FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray };
        btnSave.Click += (s, e) => { SaveConfig(); MessageBox.Show("設定已儲存，重啟程式後生效"); };
        r2.Controls.AddRange(new Control[] { cmbFreq, cmbDepth, btnSave });

        // 3. 監控列表標頭與清除按鈕
        FlowLayoutPanel r3 = new FlowLayoutPanel() { AutoSize = true, Width = 340, Margin = new Padding(0, 10, 0, 0) };
        Label lblList = new Label() { Text = "🔍 異動紀錄清單", AutoSize = true, Font = new Font(MainFont, FontStyle.Bold), ForeColor = Color.DimGray, Margin = new Padding(0, 8, 80, 0) };
        Button btnClear = new Button() { Text = "🗑️ 一鍵清除", Width = 80, Height = 25, FlatStyle = FlatStyle.Flat, BackColor = Color.White };
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();
        r3.Controls.AddRange(new Control[] { lblList, btnClear });

        mainPanel.Controls.AddRange(new Control[] { r1, r2, r3 });
        this.Controls.Add(mainPanel);

        // --- 卡片清單區 ---
        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        this.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        LoadConfigAndStartWatch();
    }

    private void AddNewPath(string src, string dst) {
        if (!Directory.Exists(src)) { MessageBox.Show("來源路徑無效！"); return; }
        if (string.IsNullOrEmpty(dst)) { MessageBox.Show("請填寫備份路徑！"); return; }
        pathPairs[src] = dst;
        SaveConfig();
        StartWatcher(src);
        txtSource.Text = ""; txtBackup.Text = "";
        MessageBox.Show("已新增監控路徑");
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            if (line.StartsWith("#") || string.IsNullOrWhiteSpace(line)) {
                // 讀取設定參數
                if (line.StartsWith("#SET|")) {
                    var p = line.Split('|');
                    if (p.Length >= 3) { cmbFreq.Text = p[1]; cmbDepth.Text = p[2]; }
                }
                continue;
            }
            string[] parts = line.Split('|');
            if (parts.Length >= 2 && Directory.Exists(parts[0].Trim())) {
                pathPairs[parts[0].Trim()] = parts[1].Trim();
                StartWatcher(parts[0].Trim());
            }
        }
    }

    private void SaveConfig() {
        List<string> lines = new List<string>();
        lines.Add(string.Format("#SET|{0}|{1}", cmbFreq.Text, cmbDepth.Text));
        lines.Add("# 格式: 來源路徑|備份路徑");
        foreach (var kvp in pathPairs) lines.Add(kvp.Key + "|" + kvp.Value);
        File.WriteAllLines(configFile, lines);
    }

    private void StartWatcher(string path) {
        FileSystemWatcher w = new FileSystemWatcher(path);
        w.IncludeSubdirectories = (cmbDepth.Text == "無限層");
        w.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName;
        w.Changed += OnFileEvent; w.Created += OnFileEvent;
        w.EnableRaisingEvents = true;
        watchers.Add(w);
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        lock (lockObj) {
            DateTime now = DateTime.Now;
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = now;
        }
        string srcRoot = ((FileSystemWatcher)sender).Path;
        if (pathPairs.ContainsKey(srcRoot)) DoBackup(srcRoot, e.FullPath, pathPairs[srcRoot]);
    }

    private void DoBackup(string srcRoot, string fileFullPath, string dstRoot) {
        try {
            Uri rootUri = new Uri(srcRoot.EndsWith("\\") ? srcRoot : srcRoot + "\\");
            string relativePath = Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(fileFullPath)).ToString().Replace('/', '\\'));
            string destFile = Path.Combine(dstRoot, relativePath);
            string destDir = Path.GetDirectoryName(destFile);

            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
            
            // 依據設定的頻率決定延遲（簡單模擬）
            System.Threading.Thread.Sleep(500);
            File.Copy(fileFullPath, destFile, true);

            this.Invoke(new Action(() => {
                Panel card = CreateCard(Path.GetFileName(fileFullPath), relativePath, fileFullPath, destFile);
                cardPanel.Controls.Add(card);
                cardPanel.Controls.SetChildIndex(card, 0);
                if (cardPanel.Controls.Count > 15) cardPanel.Controls.RemoveAt(15);
            }));
        } catch { }
    }

    private Panel CreateCard(string fileName, string relPath, string origPath, string backupPath) {
        Panel card = new Panel() { Width = 330, Height = 75, BackColor = Color.FromArgb(252, 252, 254), Margin = new Padding(0, 3, 0, 3), BorderStyle = BorderStyle.FixedSingle };
        Label lbl = new Label() { Text = string.Format("{0}\n路徑: {1}", fileName, relPath), Location = new Point(8, 8), Width = 230, Height = 45, Font = MainFont };
        Button btnView = new Button() { Text = "查看", Location = new Point(255, 12), Width = 55, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        btnView.Click += (s, e) => { try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + origPath + "\""); } catch { } };
        LinkLabel lnk = new LinkLabel() { Text = "開啟備份檔", Location = new Point(250, 48), AutoSize = true, Font = new Font(MainFont.FontFamily, 8f) };
        lnk.Click += (s, e) => { try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + backupPath + "\""); } catch { } };
        card.Controls.AddRange(new Control[] { lbl, btnView, lnk });
        return card;
    }
}
