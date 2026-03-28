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
    
    private List<string> watchPaths = new List<string>();
    private Dictionary<string, DateTime> knownFiles = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

    private NotifyIcon trayIcon;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel fileListPanel;
    private Label titleLabel;
    private HashSet<string> currentFiles = new HashSet<string>();
    
    private DateTime startTime;
    private string appName = "MacosTodoWatcherApp"; 
    private bool isPositionLocked = false; 
    private string backupDirectory = "";

    private static Color BgColor = Color.FromArgb(245, 245, 247); 
    private static Color CardColor = Color.White;
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public MacosTodoWatcher() {
        // 【關鍵修復 1】強制建立視窗神經，但不使用破壞繪圖的底層隱藏
        IntPtr forceHandle = this.Handle;
        
        startTime = DateTime.Now; 

        this.Text = "通知中心";
        this.Width = 360;
        this.Height = 450;
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
            Text = "📋 待處理項目：0", Dock = DockStyle.Top, Height = 50, 
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font(MainFont.FontFamily, 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };
        this.Controls.Add(titleLabel);

        fileListPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(12, 0, 12, 10), BackColor = BgColor 
        };
        this.Controls.Add(fileListPanel);

        // --- 系統托盤與右鍵選單 ---
        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "檔案監控 (掃描模式)" };
        ContextMenu menu = new ContextMenu();
        
        menu.MenuItems.Add("顯示待辦清單", new EventHandler(delegate { ShowAppWindow(); }));
        menu.MenuItems.Add("⚡ 強制立即掃描 (測試用)", new EventHandler(OnForceScanClick));
        menu.MenuItems.Add("-");
        
        MenuItem startupMenu = new MenuItem("✅ 開機自動執行");
        startupMenu.Checked = IsRunOnStartup(); 
        startupMenu.Click += new EventHandler(ToggleStartup);
        menu.MenuItems.Add(startupMenu);
        
        MenuItem lockMenu = new MenuItem("📌 鎖定視窗位置");
        lockMenu.Click += new EventHandler(delegate {
            isPositionLocked = !isPositionLocked;
            lockMenu.Checked = isPositionLocked;
        });
        menu.MenuItems.Add(lockMenu);
        menu.MenuItems.Add("-");

        MenuItem timerMenu = new MenuItem("⏱️ 掃描頻率設定");
        timerMenu.MenuItems.Add("即時掃描 (1 秒)", new EventHandler(delegate { SetScanInterval(1000); }));
        timerMenu.MenuItems.Add("一般掃描 (5 秒)", new EventHandler(delegate { SetScanInterval(5000); }));
        timerMenu.MenuItems.Add("自訂計時掃描...", new EventHandler(OnCustomTimerClick));
        menu.MenuItems.Add(timerMenu);
        
        MenuItem depthMenu = new MenuItem("📂 子資料夾監控深度");
        depthMenu.MenuItems.Add("僅當前資料夾 (0層)", new EventHandler(delegate { SetMaxDepth(0); }));
        depthMenu.MenuItems.Add("向下 1 層", new EventHandler(delegate { SetMaxDepth(1); }));
        depthMenu.MenuItems.Add("向下 2 層", new EventHandler(delegate { SetMaxDepth(2); }));
        depthMenu.MenuItems.Add("無限層 (預設)", new EventHandler(delegate { SetMaxDepth(-1); }));
        depthMenu.MenuItems.Add("自訂層數...", new EventHandler(OnCustomDepthClick));
        menu.MenuItems.Add(depthMenu);
        
        menu.MenuItems.Add("-");
        menu.MenuItems.Add("➕ 新增監控資料夾...", new EventHandler(OnAddFolderClick));
        menu.MenuItems.Add("📁 設定自動備份資料夾...", new EventHandler(OnSetBackupFolder));
        menu.MenuItems.Add("⚙️ 管理監控與狀態...", new EventHandler(ShowManageWindow));
        menu.MenuItems.Add("-");
        menu.MenuItems.Add("結束程式", new EventHandler(delegate { trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = menu;

        scanTimer = new Timer();
        scanTimer.Tick += new EventHandler(PerformScan);
        
        LoadConfig();
        
        scanTimer.Interval = scanIntervalMs;
        scanTimer.Start();

        // 【關鍵修復 2】安全的啟動隱藏法：先變透明，載入後隱藏並恢復透明度，確保繪圖引擎正常
        this.Opacity = 0; 
        this.Load += new EventHandler(delegate { 
            this.Hide(); 
            this.Opacity = 1; 
        });
    }

    // ==========================================
    // 獨立顯示視窗功能 (確保畫面刷新)
    // ==========================================
    private void ShowAppWindow() {
        this.Show();
        if (this.WindowState == FormWindowState.Minimized) {
            this.WindowState = FormWindowState.Normal;
        }
        this.Activate(); // 將視窗推到最上層
        this.Refresh();  // 強制重繪內容
    }

    private void OnForceScanClick(object sender, EventArgs e) {
        PerformScan(null, null); 
        
        int totalDiskFiles = 0;
        foreach (string dir in watchPaths) {
            if (Directory.Exists(dir)) { totalDiskFiles += GetFilesByDepth(dir, maxScanDepth).Count; }
        }

        string report = string.Format("掃描完畢！\n\n監控資料夾數：{0} 個\n記憶體紀錄檔案數：{1} 個\n硬碟實際看見檔案數：{2} 個\n\n※ 視窗應該已經正常彈出了。", 
            watchPaths.Count, knownFiles.Count, totalDiskFiles);
            
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
        if (watchPaths.Count == 0) return;
        if (scanTimer != null) scanTimer.Stop(); 

        HashSet<string> currentDiskFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string dir in watchPaths) {
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
            foreach (string dir in watchPaths) {
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

    private void OnFileChanged(string fullPath) {
        trayIcon.ShowBalloonTip(3000, "偵測到檔案變動", "檔案: " + Path.GetFileName(fullPath), ToolTipIcon.Info);
        AutoBackupFile(fullPath);
        SyncAdd(fullPath); 
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

    private void OnSetBackupFolder(object sender, EventArgs e) {
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        fbd.Description = "請選擇自動備份的目標資料夾：";
        if (fbd.ShowDialog() == DialogResult.OK) {
            backupDirectory = fbd.SelectedPath;
            SaveConfig();
            MessageBox.Show("備份資料夾已設定為：\n" + backupDirectory, "設定成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        fbd.Dispose();
    }

    private void SaveConfig() {
        List<string> lines = new List<string>();
        lines.Add("ScanInterval=" + scanIntervalMs.ToString());
        lines.Add("MaxDepth=" + maxScanDepth.ToString()); 
        if (!string.IsNullOrEmpty(backupDirectory)) { lines.Add("BackupDir=" + backupDirectory); }
        foreach (string path in watchPaths) { lines.Add(path); }
        File.WriteAllLines(configFile, lines.ToArray());
    }

    private void AutoBackupFile(string sourceFile) {
        if (string.IsNullOrEmpty(backupDirectory) || !Directory.Exists(backupDirectory)) return;
        string sourceDir = Path.GetDirectoryName(sourceFile);
        if (sourceDir.StartsWith(backupDirectory, StringComparison.OrdinalIgnoreCase)) return;

        System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(delegate(object state) {
            string destFile = Path.Combine(backupDirectory, Path.GetFileName(sourceFile));
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

    private void ShowManageWindow(object sender, EventArgs e) {
        Form mgr = new Form() {
            Text = "管理監控清單與狀態", Width = 450, Height = 420, 
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = false, MinimizeBox = false, BackColor = Color.White
        };
        
        TimeSpan uptime = DateTime.Now - startTime;
        string upStr = string.Format("{0}天 {1}小時 {2}分鐘", uptime.Days, uptime.Hours, uptime.Minutes);
        string bkStr = string.IsNullOrEmpty(backupDirectory) ? "尚未設定" : backupDirectory;
        string freqStr = (scanIntervalMs / 1000.0).ToString() + " 秒";
        string depthStr = maxScanDepth == -1 ? "無限層 (全部)" : "向下 " + maxScanDepth.ToString() + " 層";

        Label lblStat = new Label() {
            Text = "⏱️ 程式已運行：" + upStr + "\n📁 監控中資料夾：" + watchPaths.Count.ToString() + " 個\n💾 自動備份至：" + bkStr + "\n📡 掃描頻率：" + freqStr + "\n📂 監控深度：" + depthStr,
            Location = new Point(20, 15), AutoSize = true, Font = new Font(MainFont, FontStyle.Bold),
            ForeColor = Color.FromArgb(60, 60, 60)
        };
        mgr.Controls.Add(lblStat);

        ListBox lb = new ListBox() { Location = new Point(20, 115), Width = 390, Height = 140, Font = MainFont };
        foreach(var w in watchPaths) { lb.Items.Add(w); }
        mgr.Controls.Add(lb);

        Button btnRemove = new Button() {
            Text = "🗑️ 移除選取的資料夾", Location = new Point(20, 270), Width = 180, Height = 35,
            FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold)
        };
        btnRemove.FlatAppearance.BorderSize = 0;

        btnRemove.Click += new EventHandler(delegate {
            if(lb.SelectedIndex != -1) {
                string selPath = lb.SelectedItem.ToString();
                watchPaths.Remove(selPath);
                lb.Items.RemoveAt(lb.SelectedIndex);
                SaveConfig();
                lblStat.Text = "⏱️ 程式已運行：" + upStr + "\n📁 監控中資料夾：" + watchPaths.Count.ToString() + " 個\n💾 自動備份至：" + bkStr + "\n📡 掃描頻率：" + freqStr + "\n📂 監控深度：" + depthStr;
            } else { MessageBox.Show("請先選擇要移除的資料夾。", "提示"); }
        });
        mgr.Controls.Add(btnRemove);
        mgr.ShowDialog();
    }

    private void OnAddFolderClick(object sender, EventArgs e) {
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        fbd.Description = "請選擇要新增監控的資料夾：";
        if (fbd.ShowDialog() == DialogResult.OK) { AddNewPath(fbd.SelectedPath); }
        fbd.Dispose();
    }

    private void AddNewPath(string newPath) {
        if (!Directory.Exists(newPath)) return;
        foreach (string p in watchPaths) {
            if (p.Equals(newPath, StringComparison.OrdinalIgnoreCase)) {
                MessageBox.Show("這個資料夾已經在監控清單中了！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        watchPaths.Add(newPath); QuietlyIndexDirectory(newPath); SaveConfig();
        trayIcon.ShowBalloonTip(3000, "新增成功", "已開始掃描：\n" + newPath, ToolTipIcon.Info);
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
            if (text.StartsWith("BackupDir=", StringComparison.OrdinalIgnoreCase)) {
                backupDirectory = text.Substring(10).Trim(); continue;
            }
            if (Directory.Exists(text)) { watchPaths.Add(text); QuietlyIndexDirectory(text); }
        }
    }

    // ==========================================
    // 【關鍵修復 3】確保 UI 絕對強制渲染，並加上 ForeColor 以防顏色被吃掉
    // ==========================================
    private void SyncAdd(string path) {
        if (this.InvokeRequired) { 
            this.BeginInvoke(new Action<string>(SyncAdd), new object[] { path }); 
            return; 
        }
        if (currentFiles.Contains(path) || Directory.Exists(path)) return;
        currentFiles.Add(path);

        Panel card = new Panel() { Width = 310, Height = 60, BackColor = CardColor, Margin = new Padding(0, 0, 0, 10) };
        card.Tag = path;

        card.Paint += new PaintEventHandler(delegate(object s, PaintEventArgs ev) {
            try {
                using (GraphicsPath p = GetRoundedPath(new Rectangle(0, 0, card.Width-1, card.Height-1), 10)) {
                    ev.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    ev.Graphics.DrawPath(new Pen(Color.FromArgb(230, 230, 230)), p);
                }
            } catch {} // 確保繪製邊框就算報錯也不會導致裡面內容消失
        });

        Label name = new Label() { Text = Path.GetFileName(path), Location = new Point(15, 12), Width = 210, Font = MainFont, AutoEllipsis = true, ForeColor = Color.Black };
        Label info = new Label() { Text = DateTime.Now.ToString("HH:mm") + " • 新變動", Location = new Point(15, 34), ForeColor = Color.Gray, Font = new Font(MainFont.FontFamily, 8) };
        
        Button btn = new Button() { 
            Text = "查看", Location = new Point(235, 14), Width = 60, Height = 32, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) 
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

    private void UpdateCount() { titleLabel.Text = "📋 待處理項目：" + fileListPanel.Controls.Count.ToString(); }

    private GraphicsPath GetRoundedPath(Rectangle r, int rad) {
        GraphicsPath p = new GraphicsPath(); int d = rad * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90); p.AddArc(r.Right-d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right-d, r.Bottom-d, d, d, 0, 90); p.AddArc(r.X, r.Bottom-d, d, d, 90, 90);
        p.CloseFigure(); return p;
    }

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
