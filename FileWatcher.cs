using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using Microsoft.Win32;
using System.Runtime.InteropServices; 

public class MacosTodoWatcher : Form {
    private Timer scanTimer;
    private int scanIntervalMs = 1500; 
    private int maxScanDepth = -1; 
    
    private class WatchTask {
        public string BackupPath;
        public bool IsSilent;
        public WatchTask(string backup, bool silent) { BackupPath = backup; IsSilent = silent; }
    }
    
    private Dictionary<string, WatchTask> watchConfigs = new Dictionary<string, WatchTask>(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, DateTime> knownFiles = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

    private NotifyIcon trayIcon;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel fileListPanel;
    private Label titleLabel;
    private HashSet<string> currentFiles = new HashSet<string>();
    
    private DateTime startTime;
    private string appName = "MacosTodoWatcherApp"; 
    private bool isPositionLocked = false; 

    private static Color BgColor = Color.FromArgb(245, 245, 247); 
    private static Color CardColor = Color.White;
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    public MacosTodoWatcher() {
        IntPtr forceHandle = this.Handle;
        startTime = DateTime.Now; 

        this.Text = "通知中心";
        this.Width = 380; 
        this.Height = 420;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.MaximizeBox = true;
        this.MinimizeBox = true;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.BackColor = BgColor;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        TableLayoutPanel mainLayout = new TableLayoutPanel() {
            Dock = DockStyle.Fill,
            RowCount = 2,
            ColumnCount = 1,
            Margin = new Padding(0),
            Padding = new Padding(0)
        };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 51F));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        this.Controls.Add(mainLayout);

        Panel titleContainer = new Panel() { Dock = DockStyle.Fill, Margin = new Padding(0), BackColor = BgColor };
        titleLabel = new Label() { 
            Text = "待處理項目：0", Dock = DockStyle.Fill, 
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font(MainFont.FontFamily, 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };
        Panel titleBorder = new Panel() { Dock = DockStyle.Bottom, Height = 1, BackColor = Color.FromArgb(220, 220, 220) };
        titleContainer.Controls.Add(titleLabel);
        titleContainer.Controls.Add(titleBorder);
        
        mainLayout.Controls.Add(titleContainer, 0, 0); 

        fileListPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            AutoScroll = true, 
            Margin = new Padding(0),
            Padding = new Padding(0), 
            BackColor = BgColor 
        };
        
        mainLayout.Controls.Add(fileListPanel, 0, 1); 

        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "檔案監控 (掃描模式)" };
        ContextMenu menu = new ContextMenu();
        
        menu.MenuItems.Add("顯示待辦清單", new EventHandler(delegate { ShowAppWindow(); }));
        menu.MenuItems.Add("強制立即掃描 (測試用)", new EventHandler(OnForceScanClick));
        menu.MenuItems.Add("-");
        
        MenuItem startupMenu = new MenuItem("開機自動執行");
        startupMenu.Checked = IsRunOnStartup(); 
        startupMenu.Click += new EventHandler(ToggleStartup);
        menu.MenuItems.Add(startupMenu);
        
        MenuItem lockMenu = new MenuItem("鎖定視窗位置");
        lockMenu.Click += new EventHandler(delegate { isPositionLocked = !isPositionLocked; lockMenu.Checked = isPositionLocked; });
        menu.MenuItems.Add(lockMenu);
        menu.MenuItems.Add("-");

        // ==========================================
        // 【新增選單】擴充掃描頻率選項
        // ==========================================
        MenuItem timerMenu = new MenuItem("掃描頻率設定");
        timerMenu.MenuItems.Add("即時掃描 (1 秒)", new EventHandler(delegate { SetScanInterval(1000); }));
        timerMenu.MenuItems.Add("一般掃描 (5 秒)", new EventHandler(delegate { SetScanInterval(5000); }));
        timerMenu.MenuItems.Add("稍微放寬 (10 秒)", new EventHandler(delegate { SetScanInterval(10000); }));
        timerMenu.MenuItems.Add("省電模式 (30 秒)", new EventHandler(delegate { SetScanInterval(30000); }));
        timerMenu.MenuItems.Add("低頻模式 (1 分鐘)", new EventHandler(delegate { SetScanInterval(60000); }));
        timerMenu.MenuItems.Add("極低頻模式 (5 分鐘)", new EventHandler(delegate { SetScanInterval(300000); }));
        timerMenu.MenuItems.Add("-"); // 加入分隔線
        timerMenu.MenuItems.Add("自訂計時掃描...", new EventHandler(OnCustomTimerClick));
        menu.MenuItems.Add(timerMenu);
        // ==========================================
        
        MenuItem depthMenu = new MenuItem("子資料夾監控深度");
        depthMenu.MenuItems.Add("僅當前資料夾 (0層)", new EventHandler(delegate { SetMaxDepth(0); }));
        depthMenu.MenuItems.Add("向下 1 層", new EventHandler(delegate { SetMaxDepth(1); }));
        depthMenu.MenuItems.Add("向下 2 層", new EventHandler(delegate { SetMaxDepth(2); }));
        depthMenu.MenuItems.Add("無限層 (預設)", new EventHandler(delegate { SetMaxDepth(-1); }));
        depthMenu.MenuItems.Add("自訂層數...", new EventHandler(OnCustomDepthClick));
        menu.MenuItems.Add(depthMenu);
        
        menu.MenuItems.Add("-");
        menu.MenuItems.Add("新增監控任務 (指定來源/備份/通知)...", new EventHandler(OnAddFolderClick));
        menu.MenuItems.Add("管理監控與狀態...", new EventHandler(ShowManageWindow));
        menu.MenuItems.Add("-");
        menu.MenuItems.Add("結束程式", new EventHandler(delegate { trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = menu;

        scanTimer = new Timer();
        scanTimer.Tick += new EventHandler(PerformScan);
        
        LoadConfig();
        scanTimer.Interval = scanIntervalMs;
        scanTimer.Start();

        this.Opacity = 0; 
        this.Load += new EventHandler(delegate { this.Hide(); this.Opacity = 1; });
    }

    private void ShowAppWindow() {
        this.Show();
        if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
        this.Activate(); 
        this.Refresh();  
    }

    private void OnForceScanClick(object sender, EventArgs e) {
        PerformScan(null, null); 
        int totalDiskFiles = 0;
        foreach (string dir in watchConfigs.Keys) {
            if (Directory.Exists(dir)) { totalDiskFiles += GetFilesByDepth(dir, maxScanDepth).Count; }
        }
        MessageBox.Show(string.Format("掃描完畢！\n\n監控資料夾數：{0} 個\n記憶體紀錄檔案數：{1} 個\n硬碟實際看見檔案數：{2} 個", watchConfigs.Count, knownFiles.Count, totalDiskFiles), "強制掃描報告", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void SetScanInterval(int ms) {
        scanIntervalMs = ms;
        if (scanTimer != null) scanTimer.Interval = scanIntervalMs;
        SaveConfig();
        
        // 修正提示文字，讓超過 60 秒的顯示更直覺
        string timeStr = ms >= 60000 ? (ms / 60000.0) + " 分鐘" : (ms / 1000.0) + " 秒";
        trayIcon.ShowBalloonTip(2000, "設定更新", "掃描頻率已更改為 " + timeStr, ToolTipIcon.Info);
    }

    private void OnCustomTimerClick(object sender, EventArgs e) {
        string input = ShowInputBox("請輸入掃描間隔 (秒)：", "自訂掃描頻率", (scanIntervalMs / 1000).ToString());
        if (!string.IsNullOrEmpty(input)) {
            int sec;
            if (int.TryParse(input, out sec) && sec > 0) { SetScanInterval(sec * 1000); } 
            else { MessageBox.Show("請輸入有效的正整數數字！", "錯誤"); }
        }
    }

    private void SetMaxDepth(int depth) {
        maxScanDepth = depth;
        SaveConfig();
        trayIcon.ShowBalloonTip(2000, "深度設定更新", "子資料夾監控深度已設定為：\n" + (depth == -1 ? "無限層" : "向下 " + depth + " 層"), ToolTipIcon.Info);
    }

    private void OnCustomDepthClick(object sender, EventArgs e) {
        string input = ShowInputBox("請輸入監控深度 (整數)：\n(0 代表不看子資料夾，-1 代表無限層)", "自訂監控深度", maxScanDepth.ToString());
        if (!string.IsNullOrEmpty(input)) {
            int d;
            if (int.TryParse(input, out d) && d >= -1) { SetMaxDepth(d); } 
            else { MessageBox.Show("請輸入有效的整數數字 (>= -1)！", "錯誤"); }
        }
    }

    private class DirNode {
        public string Path; public int Depth;
        public DirNode(string p, int d) { Path = p; Depth = d; }
    }

    private List<string> GetFilesByDepth(string rootPath, int maxDepth) {
        List<string> files = new List<string>();
        Queue<DirNode> queue = new Queue<DirNode>();
        queue.Enqueue(new DirNode(rootPath, 0));

        while (queue.Count > 0) {
            DirNode current = queue.Dequeue();
            try {
                string[] dirFiles = Directory.GetFiles(current.Path);
                foreach (string f in dirFiles) { files.Add(f); }
            } catch { } 

            if (maxDepth == -1 || current.Depth < maxDepth) {
                try {
                    string[] subDirs = Directory.GetDirectories(current.Path);
                    foreach (string subDir in subDirs) { queue.Enqueue(new DirNode(subDir, current.Depth + 1)); }
                } catch { }
            }
        }
        return files;
    }

    private void QuietlyIndexDirectory(string dir) {
        if (!Directory.Exists(dir)) return;
        List<string> files = GetFilesByDepth(dir, maxScanDepth);
        foreach (string f in files) {
            try { knownFiles[f] = File.GetLastWriteTime(f); } catch {}
        }
    }

    private void PerformScan(object sender, EventArgs e) {
        if (watchConfigs.Count == 0) return;
        if (scanTimer != null) scanTimer.Stop(); 

        HashSet<string> currentDiskFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string dir in watchConfigs.Keys) {
            if (!Directory.Exists(dir)) continue;

            List<string> files = GetFilesByDepth(dir, maxScanDepth);
            foreach (string file in files) {
                currentDiskFiles.Add(file);
                try {
                    DateTime lastWrite = File.GetLastWriteTime(file);
                    bool isNewOrChanged = false;
                    
                    if (!knownFiles.ContainsKey(file)) { isNewOrChanged = true; } 
                    else if (knownFiles[file] != lastWrite) { isNewOrChanged = true; }

                    if (isNewOrChanged) {
                        knownFiles[file] = lastWrite; 
                        OnFileChanged(file); 
                    }
                } catch {}
            }
        }

        List<string> deleted = new List<string>();
        foreach (string knownPath in knownFiles.Keys) {
            bool belongsToWatch = false;
            foreach (string dir in watchConfigs.Keys) {
                if (knownPath.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) { belongsToWatch = true; break; }
            }
            if (belongsToWatch && !currentDiskFiles.Contains(knownPath)) { deleted.Add(knownPath); }
        }

        foreach (string del in deleted) { knownFiles.Remove(del); OnFileDeleted(del); }
        if (scanTimer != null) scanTimer.Start(); 
    }

    private void OnFileChanged(string fullPath) {
        AutoBackupFile(fullPath); 
        WatchTask task = GetWatchTaskForFile(fullPath);
        bool isSilent = (task != null && task.IsSilent);

        if (!isSilent) {
            trayIcon.ShowBalloonTip(3000, "偵測到檔案變動", "檔案: " + Path.GetFileName(fullPath), ToolTipIcon.Info);
            SyncAdd(fullPath); 
        }
    }
    
    private void OnFileDeleted(string fullPath) { RemoveFromFileList(fullPath); }

    public static string ShowInputBox(string prompt, string title, string defaultValue) {
        Form form = new Form() { Width = 350, Height = 175, FormBorderStyle = FormBorderStyle.FixedDialog, Text = title, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = false, MinimizeBox = false };
        Label textLabel = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true };
        TextBox textBox = new TextBox() { Left = 20, Top = 65, Width = 290, Text = defaultValue };
        Button confirmation = new Button() { Text = "確定", Left = 150, Width = 75, Top = 100, DialogResult = DialogResult.OK };
        Button cancel = new Button() { Text = "取消", Left = 235, Width = 75, Top = 100, DialogResult = DialogResult.Cancel };
        confirmation.Click += new EventHandler(delegate { form.Close(); });
        cancel.Click += new EventHandler(delegate { form.Close(); });
        form.Controls.Add(textBox); form.Controls.Add(confirmation); form.Controls.Add(cancel); form.Controls.Add(textLabel);
        form.AcceptButton = confirmation; form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK ? textBox.Text : "";
    }

    private void OnAddFolderClick(object sender, EventArgs e) {
        string sourcePath = ""; string backupPath = ""; bool isSilent = false;
        using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
            fbd.Description = "【步驟 1/3】請選擇要「監控」的資料夾：";
            if (fbd.ShowDialog() == DialogResult.OK) { sourcePath = fbd.SelectedPath; } else { return; }
        }
        if (watchConfigs.ContainsKey(sourcePath)) { MessageBox.Show("這個資料夾已經在清單中了！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

        DialogResult drBackup = MessageBox.Show("【步驟 2/3】是否設定「專屬自動備份」位置？\n\n選 [是]：指定備份位置\n選 [否]：只監控，不備份", "備份設定", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (drBackup == DialogResult.Cancel) return;
        if (drBackup == DialogResult.Yes) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                fbd.Description = "請選擇「" + Path.GetFileName(sourcePath) + "」的備份目的地：";
                if (fbd.ShowDialog() == DialogResult.OK) { backupPath = fbd.SelectedPath; } else { return; }
            }
        }

        DialogResult drNotify = MessageBox.Show("【步驟 3/3】收到新檔案時，是否要彈出「通知與清單」？\n\n選 [是]：彈出右下角通知與待辦清單 (🔔)\n選 [否]：靜音背景處理，僅執行備份 (🔇)", "通知設定", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (drNotify == DialogResult.Cancel) return;
        if (drNotify == DialogResult.No) { isSilent = true; }

        watchConfigs[sourcePath] = new WatchTask(backupPath, isSilent);
        QuietlyIndexDirectory(sourcePath); SaveConfig();
        trayIcon.ShowBalloonTip(3000, "任務新增成功", "已開始監控：" + sourcePath, ToolTipIcon.Info);
    }

    private void SaveConfig() {
        List<string> lines = new List<string>();
        lines.Add("ScanInterval=" + scanIntervalMs.ToString());
        lines.Add("MaxDepth=" + maxScanDepth.ToString()); 
        foreach (var kvp in watchConfigs) {
            string silentFlag = kvp.Value.IsSilent ? "1" : "0";
            if (string.IsNullOrEmpty(kvp.Value.BackupPath)) { lines.Add(kvp.Key + "||" + silentFlag); } 
            else { lines.Add(kvp.Key + "|" + kvp.Value.BackupPath + "|" + silentFlag); }
        }
        File.WriteAllLines(configFile, lines.ToArray());
    }

    private void LoadConfig() {
        if (!File.Exists(configFile)) return;
        string[] lines = File.ReadAllLines(configFile);
        foreach (string line in lines) {
            string text = line.Trim();
            if (text == "" || text.StartsWith("#")) continue;
            if (text.StartsWith("ScanInterval=", StringComparison.OrdinalIgnoreCase)) { int.TryParse(text.Substring(13), out scanIntervalMs); if (scanIntervalMs < 500) scanIntervalMs = 500; continue; }
            if (text.StartsWith("MaxDepth=", StringComparison.OrdinalIgnoreCase)) { int.TryParse(text.Substring(9), out maxScanDepth); continue; }
            if (text.StartsWith("BackupDir=", StringComparison.OrdinalIgnoreCase)) continue;
            
            string source = text; string backup = ""; bool isSilent = false;
            if (text.Contains("|")) {
                string[] parts = text.Split(new char[] { '|' }, 3);
                source = parts[0].Trim();
                if (parts.Length > 1) backup = parts[1].Trim();
                if (parts.Length > 2) isSilent = (parts[2].Trim() == "1");
            }
            if (Directory.Exists(source)) { watchConfigs[source] = new WatchTask(backup, isSilent); QuietlyIndexDirectory(source); }
        }
    }

    private WatchTask GetWatchTaskForFile(string fullPath) {
        string longestMatch = ""; WatchTask matchedTask = null;
        foreach (var kvp in watchConfigs) {
            if (fullPath.StartsWith(kvp.Key, StringComparison.OrdinalIgnoreCase)) {
                if (kvp.Key.Length > longestMatch.Length) { longestMatch = kvp.Key; matchedTask = kvp.Value; }
            }
        }
        return matchedTask;
    }

    private void AutoBackupFile(string sourceFile) {
        WatchTask task = GetWatchTaskForFile(sourceFile);
        if (task == null || string.IsNullOrEmpty(task.BackupPath) || !Directory.Exists(task.BackupPath)) return;
        string targetBackupDir = task.BackupPath;
        string sourceDir = Path.GetDirectoryName(sourceFile);
        if (sourceDir.StartsWith(targetBackupDir, StringComparison.OrdinalIgnoreCase)) return; 

        System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(delegate(object state) {
            string destFile = Path.Combine(targetBackupDir, Path.GetFileName(sourceFile));
            int retries = 5; 
            while (retries > 0) {
                try { System.Threading.Thread.Sleep(1000); if (File.Exists(sourceFile)) { File.Copy(sourceFile, destFile, true); } break; } 
                catch { retries--; }
            }
        }));
    }

    protected override void WndProc(ref Message m) {
        const int WM_NCLBUTTONDOWN = 0x00A1; const int HTCAPTION = 2; const int WM_SYSCOMMAND = 0x0112; const int SC_MOVE = 0xF010;          
        if (isPositionLocked) {
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt32() & 0xfff0) == SC_MOVE) return;
            if (m.Msg == WM_NCLBUTTONDOWN && m.WParam.ToInt32() == HTCAPTION) return;
        }
        base.WndProc(ref m);
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (this.WindowState == FormWindowState.Minimized) { this.Hide(); this.WindowState = FormWindowState.Normal; }
    }

    private bool IsRunOnStartup() {
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) {
                if (rk != null && rk.GetValue(appName) != null) { return rk.GetValue(appName).ToString().Replace("\"", "").Equals(Application.ExecutablePath, StringComparison.OrdinalIgnoreCase); }
            }
        } catch {} return false;
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender; item.Checked = !item.Checked;
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) {
                if (item.Checked) { rk.SetValue(appName, "\"" + Application.ExecutablePath + "\""); trayIcon.ShowBalloonTip(3000, "設定成功", "已設定開機執行！", ToolTipIcon.Info); } 
                else { rk.DeleteValue(appName, false); trayIcon.ShowBalloonTip(3000, "設定成功", "已取消開機執行。", ToolTipIcon.Info); }
            }
        } catch {}
    }

    private void ShowManageWindow(object sender, EventArgs e) {
        Form mgr = new Form() { Text = "管理監控清單", Width = 550, Height = 450, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White };
        Label lblStat = new Label() { Text = string.Format("程式已運行：{0}天 {1}小時 {2}分鐘\n監控中任務：{3} 組\n掃描頻率：{4} 秒\n監控深度：{5}", (DateTime.Now - startTime).Days, (DateTime.Now - startTime).Hours, (DateTime.Now - startTime).Minutes, watchConfigs.Count, scanIntervalMs / 1000.0, maxScanDepth == -1 ? "無限層" : maxScanDepth + "層"), Location = new Point(20, 15), AutoSize = true, Font = new Font(MainFont, FontStyle.Bold), ForeColor = Color.FromArgb(60, 60, 60) };
        mgr.Controls.Add(lblStat);

        ListBox lb = new ListBox() { Location = new Point(20, 100), Width = 490, Height = 180, Font = MainFont };
        List<string> uiKeys = new List<string>(watchConfigs.Keys);
        foreach(var key in uiKeys) { 
            WatchTask task = watchConfigs[key];
            lb.Items.Add((task.IsSilent ? "[🔇靜默] " : "[🔔通知] ") + Path.GetFileName(key) + (string.IsNullOrEmpty(task.BackupPath) ? " (無備份)" : " -> 備份: " + Path.GetFileName(task.BackupPath)));
        }
        mgr.Controls.Add(lb);

        Button btnRemove = new Button() { Text = "移除選取的任務", Location = new Point(20, 300), Width = 180, Height = 35, FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) };
        btnRemove.FlatAppearance.BorderSize = 0;
        btnRemove.Click += new EventHandler(delegate {
            if(lb.SelectedIndex != -1) { watchConfigs.Remove(uiKeys[lb.SelectedIndex]); uiKeys.RemoveAt(lb.SelectedIndex); lb.Items.RemoveAt(lb.SelectedIndex); SaveConfig(); } 
            else { MessageBox.Show("請選擇任務。", "提示"); }
        });
        mgr.Controls.Add(btnRemove);
        mgr.ShowDialog();
    }

    private void SyncAdd(string path) {
        if (string.IsNullOrWhiteSpace(path) || path.Equals(configFile, StringComparison.OrdinalIgnoreCase)) return;
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(SyncAdd), new object[] { path }); return; }
        if (currentFiles.Contains(path) || Directory.Exists(path)) return;
        currentFiles.Add(path);

        Panel card = new Panel();
        card.Size = new Size(320, 42);
        card.MinimumSize = new Size(320, 42); 
        card.BackColor = CardColor;
        
        card.Margin = new Padding(12, 10, 5, 5); 
        card.BorderStyle = BorderStyle.FixedSingle; 
        card.Tag = path;

        Label name = new Label() { Text = Path.GetFileName(path), Location = new Point(10, 11), Width = 160, Font = MainFont, AutoEllipsis = true, ForeColor = Color.Black };
        Label info = new Label() { Text = DateTime.Now.ToString("HH:mm"), Location = new Point(175, 13), AutoSize = true, ForeColor = Color.Gray, Font = new Font(MainFont.FontFamily, 8.5f) };
        Button btnView = new Button() { Text = "查看", Location = new Point(225, 7), Width = 50, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnView.FlatAppearance.BorderSize = 0;
        btnView.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 102, 204); 
        btnView.Click += new EventHandler(delegate { try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch {} RemoveFromFileList(path); });

        Button btnClose = new Button() { Text = "✕", Location = new Point(282, 7), Width = 28, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.Transparent, ForeColor = Color.DarkGray, Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnClose.FlatAppearance.BorderSize = 0;
        btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 230, 230); 
        btnClose.FlatAppearance.MouseDownBackColor = Color.FromArgb(255, 200, 200); 
        btnClose.Click += new EventHandler(delegate { RemoveFromFileList(path); }); 

        card.Controls.Add(name); 
        card.Controls.Add(info); 
        card.Controls.Add(btnView);
        card.Controls.Add(btnClose);
        
        fileListPanel.Controls.Add(card);
        fileListPanel.Controls.SetChildIndex(card, 0);

        UpdateCount();
        if (!this.Visible) { ShowAppWindow(); }
    }

    private void RemoveFromFileList(string path) {
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(RemoveFromFileList), new object[] { path }); return; }
        if (currentFiles.Contains(path)) {
            currentFiles.Remove(path);
            Control cardToRemove = null;
            foreach(Control c in fileListPanel.Controls) { if(c.Tag != null && c.Tag.ToString() == path) { cardToRemove = c; break; } }
            if(cardToRemove != null) { fileListPanel.Controls.Remove(cardToRemove); cardToRemove.Dispose(); }
            UpdateCount();
            if (fileListPanel.Controls.Count == 0) this.Hide();
        }
    }

    private void UpdateCount() { titleLabel.Text = "待處理項目：" + fileListPanel.Controls.Count.ToString(); }
    protected override void OnFormClosing(FormClosingEventArgs e) { if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); } base.OnFormClosing(e); }

    [STAThread] 
    public static void Main() { 
        if (Environment.OSVersion.Version.Major >= 6) { SetProcessDPIAware(); }
        Application.EnableVisualStyles(); 
        Application.SetCompatibleTextRenderingDefault(false); 
        Application.Run(new MacosTodoWatcher()); 
    }
}
