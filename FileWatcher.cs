using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using Microsoft.Win32;

public class MacosTodoWatcher : Form {
    private Timer scanTimer;
    private int scanIntervalMs = 1500; 
    private int maxScanDepth = -1; 
    
    // 【全新升級】建立一個專屬類別來記錄每個任務的「備份路徑」與「是否靜默」
    private class WatchTask {
        public string BackupPath;
        public bool IsSilent;
        public WatchTask(string backup, bool silent) { 
            BackupPath = backup; 
            IsSilent = silent; 
        }
    }
    
    // 將字典的值改為 WatchTask 類別
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

    public MacosTodoWatcher() {
        IntPtr forceHandle = this.Handle;
        startTime = DateTime.Now; 

        this.Text = "通知中心";
        this.Width = 345;
        this.Height = 400;
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.MaximizeBox = true;
        this.MinimizeBox = true;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.BackColor = BgColor;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        titleLabel = new Label() { 
            Text = "待處理項目：0", Dock = DockStyle.Top, Height = 50, 
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font(MainFont.FontFamily, 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };
        this.Controls.Add(titleLabel);

        fileListPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8, 0, 8, 10), BackColor = BgColor 
        };
        this.Controls.Add(fileListPanel);

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
        lockMenu.Click += new EventHandler(delegate {
            isPositionLocked = !isPositionLocked;
            lockMenu.Checked = isPositionLocked;
        });
        menu.MenuItems.Add(lockMenu);
        menu.MenuItems.Add("-");

        MenuItem timerMenu = new MenuItem("掃描頻率設定");
        timerMenu.MenuItems.Add("即時掃描 (1 秒)", new EventHandler(delegate { SetScanInterval(1000); }));
        timerMenu.MenuItems.Add("一般掃描 (5 秒)", new EventHandler(delegate { SetScanInterval(5000); }));
        timerMenu.MenuItems.Add("自訂計時掃描...", new EventHandler(OnCustomTimerClick));
        menu.MenuItems.Add(timerMenu);
        
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
        this.Load += new EventHandler(delegate { 
            this.Hide(); 
            this.Opacity = 1; 
        });
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
        string report = string.Format("掃描完畢！\n\n監控資料夾數：{0} 個\n記憶體紀錄檔案數：{1} 個\n硬碟實際看見檔案數：{2} 個", 
            watchConfigs.Count, knownFiles.Count, totalDiskFiles);
        MessageBox.Show(report, "強制掃描報告", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void SetScanInterval(int ms) {
        scanIntervalMs = ms;
        if (scanTimer != null) scanTimer.Interval = scanIntervalMs;
        SaveConfig();
        trayIcon.ShowBalloonTip(2000, "設定更新", "掃描頻率已更改為 " + (ms / 1000.0).ToString() + " 秒", ToolTipIcon.Info);
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
        string msg = depth == -1 ? "無限層 (全部子資料夾)" : "向下 " + depth.ToString() + " 層";
        trayIcon.ShowBalloonTip(2000, "深度設定更新", "子資料夾監控深度已設定為：\n" + msg, ToolTipIcon.Info);
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
        public string Path;
        public int Depth;
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
                if (knownPath.StartsWith(dir, StringComparison.OrdinalIgnoreCase)) {
                    belongsToWatch = true; break;
                }
            }
            if (belongsToWatch && !currentDiskFiles.Contains(knownPath)) { deleted.Add(knownPath); }
        }

        foreach (string del in deleted) {
            knownFiles.Remove(del);
            OnFileDeleted(del); 
        }

        if (scanTimer != null) scanTimer.Start(); 
    }

    // 【核心邏輯升級】自動根據設定決定要不要跳出通知
    private void OnFileChanged(string fullPath) {
        AutoBackupFile(fullPath); // 第一步：不管靜音與否，強制執行備份
        
        WatchTask task = GetWatchTaskForFile(fullPath);
        bool isSilent = (task != null && task.IsSilent);

        // 第二步：只有在「非靜默」的情況下，才彈出清單跟氣泡通知
        if (!isSilent) {
            trayIcon.ShowBalloonTip(3000, "偵測到檔案變動", "檔案: " + Path.GetFileName(fullPath), ToolTipIcon.Info);
            SyncAdd(fullPath); 
        }
    }
    
    private void OnFileDeleted(string fullPath) { 
        RemoveFromFileList(fullPath); 
    }

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

    // 【新增任務流程升級】加入第三步詢問通知模式
    private void OnAddFolderClick(object sender, EventArgs e) {
        string sourcePath = "";
        string backupPath = "";
        bool isSilent = false;

        using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
            fbd.Description = "【步驟 1/3】請選擇要「監控」的資料夾：";
            if (fbd.ShowDialog() == DialogResult.OK) { sourcePath = fbd.SelectedPath; } 
            else { return; }
        }

        if (watchConfigs.ContainsKey(sourcePath)) {
            MessageBox.Show("這個資料夾已經在監控清單中了！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        DialogResult drBackup = MessageBox.Show("【步驟 2/3】是否要為這個資料夾設定「專屬自動備份」位置？\n\n選 [是]：指定專屬備份位置\n選 [否]：只監控，不進行備份", "備份設定", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (drBackup == DialogResult.Cancel) return;

        if (drBackup == DialogResult.Yes) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                fbd.Description = "請選擇「" + Path.GetFileName(sourcePath) + "」的【備份】目的地：";
                if (fbd.ShowDialog() == DialogResult.OK) { backupPath = fbd.SelectedPath; } 
                else { return; }
            }
        }

        DialogResult drNotify = MessageBox.Show("【步驟 3/3】收到新檔案時，是否要彈出「通知與清單」？\n\n選 [是]：正常彈出右下角通知與待辦清單 (🔔)\n選 [否]：靜音背景處理，僅執行備份不打擾 (🔇)", "通知設定", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (drNotify == DialogResult.Cancel) return;
        
        if (drNotify == DialogResult.No) {
            isSilent = true;
        }

        watchConfigs[sourcePath] = new WatchTask(backupPath, isSilent);
        QuietlyIndexDirectory(sourcePath); 
        SaveConfig();
        
        string msg = "已開始監控：\n" + sourcePath;
        if (!string.IsNullOrEmpty(backupPath)) msg += "\n備份至：\n" + backupPath;
        msg += isSilent ? "\n\n(模式：🔇 靜音背景執行)" : "\n\n(模式：🔔 顯示通知與清單)";
        trayIcon.ShowBalloonTip(3000, "任務新增成功", msg, ToolTipIcon.Info);
    }

    // 【設定檔存取升級】寫入與讀取時加上靜音標籤 "|0" 或 "|1"
    private void SaveConfig() {
        List<string> lines = new List<string>();
        lines.Add("ScanInterval=" + scanIntervalMs.ToString());
        lines.Add("MaxDepth=" + maxScanDepth.ToString()); 
        
        foreach (var kvp in watchConfigs) {
            string silentFlag = kvp.Value.IsSilent ? "1" : "0";
            if (string.IsNullOrEmpty(kvp.Value.BackupPath)) {
                lines.Add(kvp.Key + "||" + silentFlag); 
            } else {
                lines.Add(kvp.Key + "|" + kvp.Value.BackupPath + "|" + silentFlag); 
            }
        }
        File.WriteAllLines(configFile, lines.ToArray());
    }

    private void LoadConfig() {
        if (!File.Exists(configFile)) return;
        string[] lines = File.ReadAllLines(configFile);
        foreach (string line in lines) {
            string text = line.Trim();
            if (text == "" || text.StartsWith("#")) continue;
            if (text.StartsWith("ScanInterval=", StringComparison.OrdinalIgnoreCase)) {
                int.TryParse(text.Substring(13), out scanIntervalMs);
                if (scanIntervalMs < 500) scanIntervalMs = 500; 
                continue;
            }
            if (text.StartsWith("MaxDepth=", StringComparison.OrdinalIgnoreCase)) {
                int.TryParse(text.Substring(9), out maxScanDepth); continue;
            }
            if (text.StartsWith("BackupDir=", StringComparison.OrdinalIgnoreCase)) continue;
            
            string source = text;
            string backup = "";
            bool isSilent = false;
            
            if (text.Contains("|")) {
                string[] parts = text.Split(new char[] { '|' }, 3);
                source = parts[0].Trim();
                if (parts.Length > 1) backup = parts[1].Trim();
                if (parts.Length > 2) isSilent = (parts[2].Trim() == "1");
            }

            if (Directory.Exists(source)) { 
                watchConfigs[source] = new WatchTask(backup, isSilent); 
                QuietlyIndexDirectory(source); 
            }
        }
    }

    private WatchTask GetWatchTaskForFile(string fullPath) {
        string longestMatch = "";
        WatchTask matchedTask = null;
        foreach (var kvp in watchConfigs) {
            if (fullPath.StartsWith(kvp.Key, StringComparison.OrdinalIgnoreCase)) {
                if (kvp.Key.Length > longestMatch.Length) {
                    longestMatch = kvp.Key;
                    matchedTask = kvp.Value;
                }
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
                try {
                    System.Threading.Thread.Sleep(1000); 
                    if (File.Exists(sourceFile)) { File.Copy(sourceFile, destFile, true); }
                    break; 
                } catch { retries--; }
            }
        }));
    }

    protected override void WndProc(ref Message m) {
        const int WM_NCLBUTTONDOWN = 0x00A1; 
        const int HTCAPTION = 2;             
        const int WM_SYSCOMMAND = 0x0112;    
        const int SC_MOVE = 0xF010;          

        if (isPositionLocked) {
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt32() & 0xfff0) == SC_MOVE) return;
            if (m.Msg == WM_NCLBUTTONDOWN && m.WParam.ToInt32() == HTCAPTION) return;
        }
        base.WndProc(ref m);
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        if (this.WindowState == FormWindowState.Minimized) {
            this.Hide(); 
            this.WindowState = FormWindowState.Normal; 
        }
    }

    private bool IsRunOnStartup() {
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) {
                if (rk != null && rk.GetValue(appName) != null) {
                    string regPath = rk.GetValue(appName).ToString().Replace("\"", "");
                    return regPath.Equals(Application.ExecutablePath, StringComparison.OrdinalIgnoreCase);
                }
            }
        } catch {}
        return false;
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender;
        item.Checked = !item.Checked;
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) {
                if (item.Checked) { rk.SetValue(appName, "\"" + Application.ExecutablePath + "\""); trayIcon.ShowBalloonTip(3000, "設定成功", "程式已設定為開機自動執行！", ToolTipIcon.Info); } 
                else { rk.DeleteValue(appName, false); trayIcon.ShowBalloonTip(3000, "設定成功", "已取消開機自動執行。", ToolTipIcon.Info); }
            }
        } catch (Exception ex) { MessageBox.Show("設定失敗：" + ex.Message, "錯誤"); item.Checked = !item.Checked; }
    }

    // 【管理介面升級】顯示 🔔 與 🔇 圖示
    private void ShowManageWindow(object sender, EventArgs e) {
        Form mgr = new Form() {
            Text = "管理監控清單與狀態", Width = 550, Height = 450, 
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = false, MinimizeBox = false, BackColor = Color.White
        };
        
        TimeSpan uptime = DateTime.Now - startTime;
        string upStr = string.Format("{0}天 {1}小時 {2}分鐘", uptime.Days, uptime.Hours, uptime.Minutes);
        string freqStr = (scanIntervalMs / 1000.0).ToString() + " 秒";
        string depthStr = maxScanDepth == -1 ? "無限層 (全部)" : "向下 " + maxScanDepth.ToString() + " 層";

        Label lblStat = new Label() {
            Text = "程式已運行：" + upStr + "\n監控中任務：" + watchConfigs.Count.ToString() + " 組\n掃描頻率：" + freqStr + "\n監控深度：" + depthStr,
            Location = new Point(20, 15), AutoSize = true, Font = new Font(MainFont, FontStyle.Bold),
            ForeColor = Color.FromArgb(60, 60, 60)
        };
        mgr.Controls.Add(lblStat);

        ListBox lb = new ListBox() { Location = new Point(20, 100), Width = 490, Height = 180, Font = MainFont };
        
        List<string> uiKeys = new List<string>(watchConfigs.Keys);
        foreach(var key in uiKeys) { 
            WatchTask task = watchConfigs[key];
            string modeIcon = task.IsSilent ? "[🔇靜默] " : "[🔔通知] ";
            string backupInfo = string.IsNullOrEmpty(task.BackupPath) ? "(無備份)" : "-> 備份: " + Path.GetFileName(task.BackupPath);
            lb.Items.Add(modeIcon + Path.GetFileName(key) + " " + backupInfo);
        }
        mgr.Controls.Add(lb);

        Button btnRemove = new Button() {
            Text = "移除選取的任務", Location = new Point(20, 300), Width = 180, Height = 35,
            FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold)
        };
        btnRemove.FlatAppearance.BorderSize = 0;

        btnRemove.Click += new EventHandler(delegate {
            if(lb.SelectedIndex != -1) {
                string keyToRemove = uiKeys[lb.SelectedIndex];
                watchConfigs.Remove(keyToRemove);
                uiKeys.RemoveAt(lb.SelectedIndex);
                lb.Items.RemoveAt(lb.SelectedIndex);
                SaveConfig();
                lblStat.Text = "程式已運行：" + upStr + "\n監控中任務：" + watchConfigs.Count.ToString() + " 組\n掃描頻率：" + freqStr + "\n監控深度：" + depthStr;
            } else { MessageBox.Show("請先選擇要移除的任務。", "提示"); }
        });
        mgr.Controls.Add(btnRemove);
        mgr.ShowDialog();
    }

    private void SyncAdd(string path) {
        if (string.IsNullOrWhiteSpace(path)) return;
        if (path.Equals(configFile, StringComparison.OrdinalIgnoreCase)) return;

        if (this.InvokeRequired) { 
            this.BeginInvoke(new Action<string>(SyncAdd), new object[] { path }); 
            return; 
        }
        if (currentFiles.Contains(path) || Directory.Exists(path)) return;
        currentFiles.Add(path);

        Panel card = new Panel();
        card.Size = new Size(310, 36);
        card.MinimumSize = new Size(310, 36); 
        card.BackColor = CardColor;
        card.Margin = new Padding(0, 0, 0, 4); 
        card.BorderStyle = BorderStyle.FixedSingle; 
        card.Tag = path;

        Label name = new Label() { Text = Path.GetFileName(path), Location = new Point(5, 8), Width = 190, Font = MainFont, AutoEllipsis = true, ForeColor = Color.Black };
        Label info = new Label() { Text = DateTime.Now.ToString("HH:mm"), Location = new Point(200, 10), AutoSize = true, ForeColor = Color.Gray, Font = new Font(MainFont.FontFamily, 8) };
        
        Button btn = new Button() { 
            Text = "查看", Location = new Point(245, 4), Width = 55, Height = 26, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold) 
        };
        btn.FlatAppearance.BorderSize = 0;

        btn.Click += new EventHandler(delegate {
            try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch {}
            RemoveFromFileList(path);
        });

        card.Controls.Add(name); card.Controls.Add(info); card.Controls.Add(btn);
        fileListPanel.Controls.Add(card);
        UpdateCount();
        if (!this.Visible) { ShowAppWindow(); }
    }

    private void RemoveFromFileList(string path) {
        if (this.InvokeRequired) { 
            this.BeginInvoke(new Action<string>(RemoveFromFileList), new object[] { path }); 
            return; 
        }
        if (currentFiles.Contains(path)) {
            currentFiles.Remove(path);
            Control cardToRemove = null;
            foreach(Control c in fileListPanel.Controls) {
                if(c.Tag != null && c.Tag.ToString() == path) { cardToRemove = c; break; }
            }
            if(cardToRemove != null) { fileListPanel.Controls.Remove(cardToRemove); cardToRemove.Dispose(); }
            UpdateCount();
            if (fileListPanel.Controls.Count == 0) this.Hide();
        }
    }

    private void UpdateCount() { titleLabel.Text = "待處理項目：" + fileListPanel.Controls.Count.ToString(); }

    protected override void OnFormClosing(FormClosingEventArgs e) { 
        if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); } 
        base.OnFormClosing(e); 
    }

    [STAThread] 
    public static void Main() { 
        Application.EnableVisualStyles(); 
        Application.Run(new MacosTodoWatcher()); 
    }
}
