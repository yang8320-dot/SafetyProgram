using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using Microsoft.Win32; // 【新增】用來處理開機自動執行

public class MacosTodoWatcher : Form {
    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private NotifyIcon trayIcon;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel fileListPanel;
    private Label titleLabel;
    private HashSet<string> currentFiles = new HashSet<string>();
    
    private DateTime startTime;
    private string appName = "MacosTodoWatcherApp"; // 登錄檔使用的名稱

    private static Color BgColor = Color.FromArgb(245, 245, 247); 
    private static Color CardColor = Color.White;
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public MacosTodoWatcher() {
        startTime = DateTime.Now; 

        this.Text = "通知中心";
        this.Width = 360;
        this.Height = 450;
        this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
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
        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "檔案監控 (macOS Style)" };
        ContextMenu menu = new ContextMenu();
        
        menu.MenuItems.Add("顯示待辦清單", new EventHandler(delegate { this.Show(); this.WindowState = FormWindowState.Normal; }));
        menu.MenuItems.Add("-");
        
        // 【新增】開機自動執行的開關
        MenuItem startupMenu = new MenuItem("開機自動執行");
        startupMenu.Checked = IsRunOnStartup(); // 程式啟動時自動檢查狀態
        startupMenu.Click += new EventHandler(ToggleStartup);
        menu.MenuItems.Add(startupMenu);
        
        menu.MenuItems.Add("-");
        menu.MenuItems.Add("➕ 新增監控資料夾...", new EventHandler(OnAddFolderClick));
        menu.MenuItems.Add("⚙️ 管理監控與狀態...", new EventHandler(ShowManageWindow));
        menu.MenuItems.Add("-");
        menu.MenuItems.Add("結束程式", new EventHandler(delegate { trayIcon.Dispose(); Application.Exit(); }));
        trayIcon.ContextMenu = menu;

        LoadConfig();
        
        this.Opacity = 0; 
        this.Load += new EventHandler(delegate { this.Hide(); });
    }

    // ==========================================
    // 【新增】開機自動執行邏輯
    // ==========================================
    private bool IsRunOnStartup() {
        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) {
                if (rk != null) {
                    object value = rk.GetValue(appName);
                    if (value != null) {
                        // 檢查登錄檔內的路徑是否跟現在程式的路徑一樣 (避免移動檔案後失效)
                        string regPath = value.ToString().Replace("\"", "");
                        return regPath.Equals(Application.ExecutablePath, StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
        } catch {}
        return false;
    }

    private void ToggleStartup(object sender, EventArgs e) {
        MenuItem item = (MenuItem)sender;
        item.Checked = !item.Checked; // 切換打勾狀態

        try {
            using (RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) {
                if (item.Checked) {
                    // 加上引號以防路徑有空白
                    rk.SetValue(appName, "\"" + Application.ExecutablePath + "\"");
                    trayIcon.ShowBalloonTip(3000, "設定成功", "程式已設定為開機自動執行！", ToolTipIcon.Info);
                } else {
                    rk.DeleteValue(appName, false);
                    trayIcon.ShowBalloonTip(3000, "設定成功", "已取消開機自動執行。", ToolTipIcon.Info);
                }
            }
        } catch (Exception ex) {
            MessageBox.Show("設定開機啟動失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            item.Checked = !item.Checked; // 如果失敗，把勾勾復原
        }
    }

    // ==========================================
    // 管理狀態面板與移除功能
    // ==========================================
    private void ShowManageWindow(object sender, EventArgs e) {
        Form mgr = new Form() {
            Text = "管理監控清單與狀態", Width = 450, Height = 350,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = false, MinimizeBox = false, BackColor = Color.White
        };
        
        TimeSpan uptime = DateTime.Now - startTime;
        string upStr = string.Format("{0}天 {1}小時 {2}分鐘", uptime.Days, uptime.Hours, uptime.Minutes);
        
        Label lblStat = new Label() {
            Text = "⏱️ 程式已運行：" + upStr + "\n📁 當前監控數量：" + watchers.Count.ToString() + " 個資料夾",
            Location = new Point(20, 15), AutoSize = true, Font = new Font(MainFont, FontStyle.Bold),
            ForeColor = Color.FromArgb(60, 60, 60)
        };
        mgr.Controls.Add(lblStat);

        ListBox lb = new ListBox() { Location = new Point(20, 65), Width = 390, Height = 170, Font = MainFont };
        foreach(var w in watchers) { lb.Items.Add(w.Path); }
        mgr.Controls.Add(lb);

        Button btnRemove = new Button() {
            Text = "🗑️ 移除選取的資料夾", Location = new Point(20, 250), Width = 180, Height = 35,
            FlatStyle = FlatStyle.Flat, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold)
        };
        btnRemove.FlatAppearance.BorderSize = 0;

        btnRemove.Click += new EventHandler(delegate {
            if(lb.SelectedIndex != -1) {
                string selPath = lb.SelectedItem.ToString();
                RemoveWatchPath(selPath);
                lb.Items.RemoveAt(lb.SelectedIndex);
                lblStat.Text = "⏱️ 程式已運行：" + upStr + "\n📁 當前監控數量：" + watchers.Count.ToString() + " 個資料夾";
            } else {
                MessageBox.Show("請先選擇要移除的資料夾。", "提示");
            }
        });
        mgr.Controls.Add(btnRemove);

        mgr.ShowDialog();
    }

    private void RemoveWatchPath(string pathToRemove) {
        FileSystemWatcher toRemove = null;
        foreach(var w in watchers) {
            if(w.Path.Equals(pathToRemove, StringComparison.OrdinalIgnoreCase)) { toRemove = w; break; }
        }
        if(toRemove != null) {
            toRemove.EnableRaisingEvents = false;
            toRemove.Dispose();
            watchers.Remove(toRemove);
        }
        if(File.Exists(configFile)) {
            string[] lines = File.ReadAllLines(configFile);
            List<string> newLines = new List<string>();
            foreach(string l in lines) {
                if(!l.Trim().Equals(pathToRemove, StringComparison.OrdinalIgnoreCase)) { newLines.Add(l); }
            }
            File.WriteAllLines(configFile, newLines.ToArray());
        }
        trayIcon.ShowBalloonTip(3000, "移除成功", "已停止監控：\n" + pathToRemove, ToolTipIcon.Info);
    }

    // ==========================================
    // 新增監控資料夾
    // ==========================================
    private void OnAddFolderClick(object sender, EventArgs e) {
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        fbd.Description = "請選擇要新增監控的資料夾：";
        if (fbd.ShowDialog() == DialogResult.OK) { AddNewPath(fbd.SelectedPath); }
        fbd.Dispose();
    }

    private void AddNewPath(string newPath) {
        if (!Directory.Exists(newPath)) return;
        if (File.Exists(configFile)) {
            string[] lines = File.ReadAllLines(configFile);
            foreach (string line in lines) {
                if (line.Trim().Equals(newPath, StringComparison.OrdinalIgnoreCase)) {
                    MessageBox.Show("這個資料夾已經在監控清單中了！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }
        try {
            string prefix = File.Exists(configFile) ? "\r\n" : "";
            File.AppendAllText(configFile, prefix + newPath);
        } catch {}

        try {
            FileSystemWatcher w = new FileSystemWatcher(newPath);
            w.Filter = "*.*";
            w.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            w.Created += new FileSystemEventHandler(OnFileEvent);
            w.Changed += new FileSystemEventHandler(OnFileEvent);
            w.Deleted += new FileSystemEventHandler(OnFileDeleted);
            w.EnableRaisingEvents = true;
            watchers.Add(w);
            trayIcon.ShowBalloonTip(3000, "新增成功", "已開始監控：\n" + newPath, ToolTipIcon.Info);
        } catch {}
    }

    // ==========================================
    // 系統初始化與事件綁定
    // ==========================================
    private void LoadConfig() {
        if (!File.Exists(configFile)) {
            File.WriteAllText(configFile, "# 請透過右下角圖示右鍵「新增監控資料夾」\r\n");
            return;
        }

        string[] lines = File.ReadAllLines(configFile);
        foreach (string line in lines) {
            string path = line.Trim();
            if (path == "" || path.StartsWith("#") || !Directory.Exists(path)) continue;
            try {
                FileSystemWatcher w = new FileSystemWatcher(path);
                w.Filter = "*.*";
                w.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
                w.Created += new FileSystemEventHandler(OnFileEvent);
                w.Changed += new FileSystemEventHandler(OnFileEvent);
                w.Deleted += new FileSystemEventHandler(OnFileDeleted);
                w.EnableRaisingEvents = true;
                watchers.Add(w);
            } catch {}
        }
    }

    private void OnFileEvent(object source, FileSystemEventArgs e) { SyncAdd(e.FullPath); }
    private void OnFileDeleted(object source, FileSystemEventArgs e) { RemoveFromFileList(e.FullPath); }

    // ==========================================
    // UI 面板操作
    // ==========================================
    private void SyncAdd(string path) {
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(SyncAdd), path); return; }
        if (currentFiles.Contains(path) || Directory.Exists(path)) return;
        currentFiles.Add(path);

        Panel card = new Panel() { Width = 310, Height = 60, BackColor = CardColor, Margin = new Padding(0, 0, 0, 10) };
        card.Tag = path;

        card.Paint += new PaintEventHandler(delegate(object s, PaintEventArgs ev) {
            using (GraphicsPath p = GetRoundedPath(new Rectangle(0, 0, card.Width-1, card.Height-1), 10)) {
                ev.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                ev.Graphics.DrawPath(new Pen(Color.FromArgb(230, 230, 230)), p);
            }
        });

        Label name = new Label() { Text = Path.GetFileName(path), Location = new Point(15, 12), Width = 210, Font = MainFont, AutoEllipsis = true };
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
        if (!this.Visible) { this.Opacity = 1; this.Show(); }
    }

    private void RemoveFromFileList(string path) {
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(RemoveFromFileList), path); return; }
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
