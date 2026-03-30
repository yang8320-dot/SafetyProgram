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

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9f);

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(8);

        Label lblHint = new Label() { Text = "📢 檔案變動將自動依目錄結構備份", Dock = DockStyle.Top, Height = 25, Font = new Font(MainFont, FontStyle.Bold), ForeColor = Color.Gray };
        this.Controls.Add(lblHint);

        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        this.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        LoadConfigAndStartWatch();
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) { File.WriteAllLines(configFile, new string[] { "C:\\Source|C:\\Backup" }); return; }
        foreach (string line in File.ReadAllLines(configFile)) {
            string[] parts = line.Split('|');
            if (parts.Length >= 2 && Directory.Exists(parts[0].Trim())) {
                pathPairs[parts[0].Trim()] = parts[1].Trim();
                StartWatcher(parts[0].Trim());
            }
        }
    }

    private void StartWatcher(string path) {
        FileSystemWatcher w = new FileSystemWatcher(path) { IncludeSubdirectories = true, EnableRaisingEvents = true };
        w.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName;
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
            System.Threading.Thread.Sleep(500);
            File.Copy(fileFullPath, destFile, true);

            this.Invoke(new Action(() => AddCard(Path.GetFileName(fileFullPath), relativePath, fileFullPath, destFile)));
        } catch { }
    }

    private void AddCard(string fileName, string relPath, string origPath, string backupPath) {
        Panel card = new Panel() { Width = 330, Height = 75, BackColor = Color.FromArgb(252, 252, 254), Margin = new Padding(0, 3, 0, 3), BorderStyle = BorderStyle.FixedSingle };
        Label lbl = new Label() { Text = string.Format("{0}\n路徑: {1}", fileName, relPath), Location = new Point(8, 8), Width = 230, Height = 45, Font = MainFont };
        
        Button btnView = new Button() { Text = "查看", Location = new Point(255, 12), Width = 55, Height = 30, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        btnView.Click += (s, e) => { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + origPath + "\""); };

        LinkLabel lnk = new LinkLabel() { Text = "開啟備份檔", Location = new Point(250, 48), AutoSize = true, Font = new Font(MainFont.FontFamily, 8f) };
        lnk.Click += (s, e) => { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + backupPath + "\""); };

        card.Controls.AddRange(new Control[] { lbl, btnView, lnk });
        cardPanel.Controls.Add(card); cardPanel.Controls.SetChildIndex(card, 0);
        if (cardPanel.Controls.Count > 15) cardPanel.Controls.RemoveAt(15);
    }
}
