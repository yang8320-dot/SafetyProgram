using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;

public class App_FileWatcher : UserControl {
    private Timer scanTimer;
    private int scanIntervalMs = 1500; 
    private int maxScanDepth = -1; 
    private MainForm parentForm;
    
    private class WatchTask {
        public string BackupPath; public bool IsSilent;
        public WatchTask(string backup, bool silent) { BackupPath = backup; IsSilent = silent; }
    }
    
    private Dictionary<string, WatchTask> watchConfigs = new Dictionary<string, WatchTask>(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, DateTime> knownFiles = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    
    private FlowLayoutPanel fileListPanel;
    private Label titleLabel;
    private HashSet<string> currentFiles = new HashSet<string>();
    private DateTime startTime;

    private static Color CardColor = Color.White;
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public App_FileWatcher(MainForm mainForm, ContextMenu trayMenu) {
        this.parentForm = mainForm;
        startTime = DateTime.Now;

        titleLabel = new Label() { Text = "待處理項目：0", Dock = DockStyle.Top, Height = 40, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(MainFont.FontFamily, 10.5f, FontStyle.Bold), ForeColor = Color.FromArgb(50, 50, 50) };
        this.Controls.Add(titleLabel);

        fileListPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8, 0, 8, 10), BackColor = Color.Transparent };
        this.Controls.Add(fileListPanel);
        fileListPanel.BringToFront();

        // --- 注入模組專屬選單 ---
        MenuItem moduleMenu = new MenuItem("📂 檔案監控設定");
        
        moduleMenu.MenuItems.Add("新增監控任務 (指定來源/備份/通知)...", new EventHandler(OnAddFolderClick));
        moduleMenu.MenuItems.Add("管理監控與狀態...", new EventHandler(ShowManageWindow));
        moduleMenu.MenuItems.Add("-");
        
        MenuItem timerMenu = new MenuItem("掃描頻率設定");
        timerMenu.MenuItems.Add("即時掃描 (1 秒)", new EventHandler(delegate { SetScanInterval(1000); }));
        timerMenu.MenuItems.Add("一般掃描 (5 秒)", new EventHandler(delegate { SetScanInterval(5000); }));
        timerMenu.MenuItems.Add("自訂計時掃描...", new EventHandler(OnCustomTimerClick));
        moduleMenu.MenuItems.Add(timerMenu);
        
        MenuItem depthMenu = new MenuItem("子資料夾監控深度");
        depthMenu.MenuItems.Add("僅當前資料夾 (0層)", new EventHandler(delegate { SetMaxDepth(0); }));
        depthMenu.MenuItems.Add("向下 1 層", new EventHandler(delegate { SetMaxDepth(1); }));
        depthMenu.MenuItems.Add("向下 2 層", new EventHandler(delegate { SetMaxDepth(2); }));
        depthMenu.MenuItems.Add("無限層 (預設)", new EventHandler(delegate { SetMaxDepth(-1); }));
        moduleMenu.MenuItems.Add(depthMenu);
        
        moduleMenu.MenuItems.Add("-");
        moduleMenu.MenuItems.Add("⚡ 強制立即掃描 (測試用)", new EventHandler(OnForceScanClick));
        
        trayMenu.MenuItems.Add(moduleMenu);

        scanTimer = new Timer();
        scanTimer.Tick += new EventHandler(PerformScan);
        LoadConfig();
        scanTimer.Interval = scanIntervalMs;
        scanTimer.Start();
    }

    private void SetScanInterval(int ms) { scanIntervalMs = ms; scanTimer.Interval = ms; SaveConfig(); parentForm.trayIcon.ShowBalloonTip(2000, "設定更新", "掃描頻率已更改為 " + (ms / 1000.0) + " 秒", ToolTipIcon.Info); }
    private void OnCustomTimerClick(object sender, EventArgs e) {
        string input = ShowInputBox("請輸入掃描間隔 (秒)：", "自訂掃描頻率", (scanIntervalMs / 1000).ToString());
        if (!string.IsNullOrEmpty(input)) { int sec; if (int.TryParse(input, out sec) && sec > 0) SetScanInterval(sec * 1000); }
    }
    private void SetMaxDepth(int depth) { maxScanDepth = depth; SaveConfig(); parentForm.trayIcon.ShowBalloonTip(2000, "設定更新", "監控深度已設定", ToolTipIcon.Info); }

    private void OnForceScanClick(object sender, EventArgs e) {
        PerformScan(null, null); 
        MessageBox.Show(string.Format("掃描完畢！\n任務數：{0} \n紀錄檔案數：{1}", watchConfigs.Count, knownFiles.Count), "報告", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private class DirNode { public string Path; public int Depth; public DirNode(string p, int d) { Path = p; Depth = d; } }

    private List<string> GetFilesByDepth(string rootPath, int maxDepth) {
        List<string> files = new List<string>(); Queue<DirNode> queue = new Queue<DirNode>(); queue.Enqueue(new DirNode(rootPath, 0));
        while (queue.Count > 0) {
            DirNode current = queue.Dequeue();
            try { foreach (string f in Directory.GetFiles(current.Path)) files.Add(f); } catch { } 
            if (maxDepth == -1 || current.Depth < maxDepth) { try { foreach (string subDir in Directory.GetDirectories(current.Path)) queue.Enqueue(new DirNode(subDir, current.Depth + 1)); } catch { } }
        } return files;
    }

    private void QuietlyIndexDirectory(string dir) {
        if (!Directory.Exists(dir)) return;
        foreach (string f in GetFilesByDepth(dir, maxScanDepth)) { try { knownFiles[f] = File.GetLastWriteTime(f); } catch {} }
    }

    private void PerformScan(object sender, EventArgs e) {
        if (watchConfigs.Count == 0) return;
        scanTimer.Stop(); 
        HashSet<string> currentDiskFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string dir in watchConfigs.Keys) {
            if (!Directory.Exists(dir)) continue;
            foreach (string file in GetFilesByDepth(dir, maxScanDepth)) {
                currentDiskFiles.Add(file);
                try {
                    DateTime lastWrite = File.GetLastWriteTime(file);
                    bool isNewOrChanged = false;
                    if (!knownFiles.ContainsKey(file) || knownFiles[file] != lastWrite) isNewOrChanged = true;
                    if (isNewOrChanged) { knownFiles[file] = lastWrite; OnFileChanged(file); }
                } catch {}
            }
        }
        List<string> deleted = new List<string>();
        foreach (string knownPath in knownFiles.Keys) {
            bool belongsToWatch = false;
            foreach (string dir in watchConfigs.Keys) { if (knownPath.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) { belongsToWatch = true; break; } }
            if (belongsToWatch && !currentDiskFiles.Contains(knownPath)) deleted.Add(knownPath);
        }
        foreach (string del in deleted) { knownFiles.Remove(del); RemoveFromFileList(del); }
        scanTimer.Start(); 
    }

    private void OnFileChanged(string fullPath) {
        AutoBackupFile(fullPath); 
        WatchTask task = GetWatchTaskForFile(fullPath);
        if (task == null || !task.IsSilent) {
            parentForm.trayIcon.ShowBalloonTip(3000, "偵測到檔案變動", Path.GetFileName(fullPath), ToolTipIcon.Info);
            SyncAdd(fullPath); 
        }
    }

    public static string ShowInputBox(string prompt, string title, string defaultValue) {
        Form form = new Form() { Width = 350, Height = 175, FormBorderStyle = FormBorderStyle.FixedDialog, Text = title, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = false, MinimizeBox = false };
        Label textLabel = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true };
        TextBox textBox = new TextBox() { Left = 20, Top = 65, Width = 290, Text = defaultValue };
        Button confirmation = new Button() { Text = "確定", Left = 150, Width = 75, Top = 100, DialogResult = DialogResult.OK };
        Button cancel = new Button() { Text = "取消", Left = 235, Width = 75, Top = 100, DialogResult = DialogResult.Cancel };
        confirmation.Click += new EventHandler(delegate { form.Close(); }); cancel.Click += new EventHandler(delegate { form.Close(); });
        form.Controls.Add(textBox); form.Controls.Add(confirmation); form.Controls.Add(cancel); form.Controls.Add(textLabel);
        form.AcceptButton = confirmation; form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK ? textBox.Text : "";
    }

    private void OnAddFolderClick(object sender, EventArgs e) {
        string sourcePath = ""; string backupPath = ""; bool isSilent = false;
        using (FolderBrowserDialog fbd = new FolderBrowserDialog()) { fbd.Description = "請選擇「監控」資料夾："; if (fbd.ShowDialog() == DialogResult.OK) sourcePath = fbd.SelectedPath; else return; }
        if (watchConfigs.ContainsKey(sourcePath)) { MessageBox.Show("已在清單中！"); return; }
        if (MessageBox.Show("是否設定「專屬備份」？", "備份", MessageBoxButtons.YesNo) == DialogResult.Yes) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) { fbd.Description = "請選擇備份目的地："; if (fbd.ShowDialog() == DialogResult.OK) backupPath = fbd.SelectedPath; else return; }
        }
        if (MessageBox.Show("收到新檔案是否靜音背景處理 (🔇)？\n選[否]則為通知模式(🔔)", "通知", MessageBoxButtons.YesNo) == DialogResult.Yes) isSilent = true;

        watchConfigs[sourcePath] = new WatchTask(backupPath, isSilent);
        QuietlyIndexDirectory(sourcePath); SaveConfig();
        parentForm.trayIcon.ShowBalloonTip(3000, "新增成功", "已監控：" + sourcePath, ToolTipIcon.Info);
    }

    private void SaveConfig() {
        List<string> lines = new List<string>();
        lines.Add("ScanInterval=" + scanIntervalMs.ToString()); lines.Add("MaxDepth=" + maxScanDepth.ToString()); 
        foreach (var kvp in watchConfigs) {
            string silentFlag = kvp.Value.IsSilent ? "1" : "0";
            lines.Add(kvp.Key + "|" + (kvp.Value.BackupPath ?? "") + "|" + silentFlag);
        }
        File.WriteAllLines(configFile, lines.ToArray());
    }

    private void LoadConfig() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            string text = line.Trim(); if (text == "" || text.StartsWith("#") || text.StartsWith("BackupDir=")) continue;
            if (text.StartsWith("ScanInterval=", StringComparison.OrdinalIgnoreCase)) { int.TryParse(text.Substring(13), out scanIntervalMs); continue; }
            if (text.StartsWith("MaxDepth=", StringComparison.OrdinalIgnoreCase)) { int.TryParse(text.Substring(9), out maxScanDepth); continue; }
            string[] parts = text.Split(new char[] { '|' }, 3);
            if (Directory.Exists(parts[0])) { watchConfigs[parts[0]] = new WatchTask(parts.Length > 1 ? parts[1] : "", parts.Length > 2 && parts[2] == "1"); QuietlyIndexDirectory(parts[0]); }
        }
    }

    private WatchTask GetWatchTaskForFile(string fullPath) {
        string longest = ""; WatchTask matched = null;
        foreach (var kvp in watchConfigs) { if (fullPath.StartsWith(kvp.Key, StringComparison.OrdinalIgnoreCase) && kvp.Key.Length > longest.Length) { longest = kvp.Key; matched = kvp.Value; } }
        return matched;
    }

    private void AutoBackupFile(string sourceFile) {
        WatchTask task = GetWatchTaskForFile(sourceFile);
        if (task == null || string.IsNullOrEmpty(task.BackupPath) || !Directory.Exists(task.BackupPath)) return;
        if (Path.GetDirectoryName(sourceFile).StartsWith(task.BackupPath, StringComparison.OrdinalIgnoreCase)) return; 
        System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(delegate(object state) {
            string dest = Path.Combine(task.BackupPath, Path.GetFileName(sourceFile));
            int r = 5; while (r > 0) { try { System.Threading.Thread.Sleep(1000); if (File.Exists(sourceFile)) File.Copy(sourceFile, dest, true); break; } catch { r--; } }
        }));
    }

    private void ShowManageWindow(object sender, EventArgs e) {
        Form mgr = new Form() { Text = "檔案監控管理", Width = 550, Height = 450, StartPosition = FormStartPosition.CenterScreen, BackColor = Color.White };
        ListBox lb = new ListBox() { Location = new Point(20, 20), Width = 490, Height = 300, Font = MainFont };
        List<string> uiKeys = new List<string>(watchConfigs.Keys);
        foreach(var key in uiKeys) lb.Items.Add((watchConfigs[key].IsSilent ? "[🔇] " : "[🔔] ") + Path.GetFileName(key) + " -> " + (string.IsNullOrEmpty(watchConfigs[key].BackupPath) ? "(無)" : Path.GetFileName(watchConfigs[key].BackupPath)));
        mgr.Controls.Add(lb);
        Button btn = new Button() { Text = "移除選取", Location = new Point(20, 340), Width = 150, Height = 35, BackColor = Color.IndianRed, ForeColor = Color.White };
        btn.Click += new EventHandler(delegate { if(lb.SelectedIndex != -1) { watchConfigs.Remove(uiKeys[lb.SelectedIndex]); uiKeys.RemoveAt(lb.SelectedIndex); lb.Items.RemoveAt(lb.SelectedIndex); SaveConfig(); } });
        mgr.Controls.Add(btn); mgr.ShowDialog();
    }

    private void SyncAdd(string path) {
        if (string.IsNullOrWhiteSpace(path) || path.Equals(configFile, StringComparison.OrdinalIgnoreCase)) return;
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(SyncAdd), new object[] { path }); return; }
        if (currentFiles.Contains(path) || Directory.Exists(path)) return;
        currentFiles.Add(path);

        Panel card = new Panel() { Size = new Size(320, 36), MinimumSize = new Size(320, 36), BackColor = CardColor, Margin = new Padding(0, 0, 0, 4), BorderStyle = BorderStyle.FixedSingle, Tag = path };
        Label name = new Label() { Text = Path.GetFileName(path), Location = new Point(10, 8), Width = 170, Font = MainFont, AutoEllipsis = true };
        Label info = new Label() { Text = DateTime.Now.ToString("HH:mm"), Location = new Point(185, 10), AutoSize = true, ForeColor = Color.Gray, Font = new Font(MainFont.FontFamily, 8) };
        Button btnView = new Button() { Text = "查看", Location = new Point(225, 4), Width = 50, Height = 26, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
        btnView.FlatAppearance.BorderSize = 0; btnView.Click += new EventHandler(delegate { try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch {} RemoveFromFileList(path); });
        Button btnClose = new Button() { Text = "✕", Location = new Point(280, 4), Width = 28, Height = 26, FlatStyle = FlatStyle.Flat, ForeColor = Color.DarkGray, Cursor = Cursors.Hand };
        btnClose.FlatAppearance.BorderSize = 0; btnClose.Click += new EventHandler(delegate { RemoveFromFileList(path); });
        
        card.Controls.Add(name); card.Controls.Add(info); card.Controls.Add(btnView); card.Controls.Add(btnClose);
        fileListPanel.Controls.Add(card); fileListPanel.Controls.SetChildIndex(card, 0);
        UpdateCount(); if (!parentForm.Visible) parentForm.ShowAppWindow();
    }

    private void RemoveFromFileList(string path) {
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(RemoveFromFileList), new object[] { path }); return; }
        if (currentFiles.Contains(path)) {
            currentFiles.Remove(path);
            Control target = null; foreach(Control c in fileListPanel.Controls) { if(c.Tag != null && c.Tag.ToString() == path) { target = c; break; } }
            if(target != null) { fileListPanel.Controls.Remove(target); target.Dispose(); }
            UpdateCount(); if (fileListPanel.Controls.Count == 0) parentForm.Hide();
        }
    }
    private void UpdateCount() { titleLabel.Text = "待處理項目：" + fileListPanel.Controls.Count.ToString(); }
}
