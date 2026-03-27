using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;

public class MacosTodoWatcher : Form {
    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private NotifyIcon trayIcon;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel fileListPanel;
    private Label titleLabel;
    private HashSet<string> currentFiles = new HashSet<string>();

    // Apple 設計規範配色
    private static Color BgColor = Color.FromArgb(245, 245, 247); 
    private static Color CardColor = Color.White;
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public MacosTodoWatcher() {
        // 視窗基本設定
        this.Text = "通知中心";
        this.Width = 360;
        this.Height = 450;
        this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.BackColor = BgColor;

        // 定位在右下角
        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(area.Right - this.Width - 15, area.Bottom - this.Height - 15);

        // 標題區
        titleLabel = new Label() { 
            Text = "📋 待處理項目：0", Dock = DockStyle.Top, Height = 50, 
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font(MainFont.FontFamily, 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(50, 50, 50)
        };
        this.Controls.Add(titleLabel);

        // 清單容器
        fileListPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(12, 0, 12, 10), BackColor = BgColor 
        };
        this.Controls.Add(fileListPanel);

        // 右下角托盤圖示
        trayIcon = new NotifyIcon() { Icon = SystemIcons.Information, Visible = true, Text = "檔案監控 (macOS Style)" };
        ContextMenu menu = new ContextMenu();
        menu.MenuItems.Add("顯示清單", (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; });
        menu.MenuItems.Add("結束程式", (s, e) => { trayIcon.Dispose(); Application.Exit(); });
        trayIcon.ContextMenu = menu;

        LoadConfig();
        this.Opacity = 0; 
        Load += (s, e) => this.Hide();
    }

    private void LoadConfig() {
        if (!File.Exists(configFile)) return;
        foreach (var path in File.ReadAllLines(configFile).Select(l => l.Trim()).Where(l => Directory.Exists(l))) {
            var w = new FileSystemWatcher(path) { Filter = "*.*", EnableRaisingEvents = true };
            w.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            w.Created += (s, e) => SyncAdd(e.FullPath);
            w.Changed += (s, e) => SyncAdd(e.FullPath);
            watchers.Add(w);
        }
    }

    private void SyncAdd(string path) {
        if (this.InvokeRequired) { this.BeginInvoke((MethodInvoker)(() => SyncAdd(path))); return; }
        if (currentFiles.Contains(path) || Directory.Exists(path)) return;
        currentFiles.Add(path);

        // 卡片容器 (圓角模擬)
        Panel card = new Panel() { Width = 310, Height = 60, BackColor = CardColor, Margin = new Padding(0, 0, 0, 10) };
        card.Paint += (s, e) => {
            using (GraphicsPath p = GetRoundedPath(new Rectangle(0, 0, card.Width-1, card.Height-1), 10))
            { e.Graphics.SmoothingMode = SmoothingMode.AntiAlias; e.Graphics.DrawPath(new Pen(Color.FromArgb(230, 230, 230)), p); }
        };

        Label name = new Label() { Text = Path.GetFileName(path), Location = new Point(15, 12), Width = 210, Font = MainFont, AutoEllipsis = true };
        Label info = new Label() { Text = DateTime.Now.ToString("HH:mm") + " • 新變動", Location = new Point(15, 34), ForeColor = Color.Gray, Font = new Font(MainFont.FontFamily, 8) };
        
        Button btn = new Button() { 
            Text = "查看", Location = new Point(235, 14), Width = 60, Height = 32, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) 
        };
        btn.FlatAppearance.BorderSize = 0;

        btn.Click += (s, e) => {
            try { Process.Start("explorer.exe", $"/select,\"{path}\""); } catch {}
            fileListPanel.Controls.Remove(card);
            currentFiles.Remove(path);
            if (fileListPanel.Controls.Count == 0) this.Hide();
            UpdateCount();
        };

        card.Controls.Add(name); card.Controls.Add(info); card.Controls.Add(btn);
        fileListPanel.Controls.Add(card);
        UpdateCount();
        if (!this.Visible) { this.Opacity = 1; this.Show(); }
    }

    private void UpdateCount() { titleLabel.Text = $"📋 待處理項目：{fileListPanel.Controls.Count}"; }
    private GraphicsPath GetRoundedPath(Rectangle r, int rad) {
        GraphicsPath p = new GraphicsPath(); int d = rad * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90); p.AddArc(r.Right-d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right-d, r.Bottom-d, d, d, 0, 90); p.AddArc(r.X, r.Bottom-d, d, d, 90, 90);
        p.CloseFigure(); return p;
    }
    protected override void OnFormClosing(FormClosingEventArgs e) { if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); } base.OnFormClosing(e); }
    [STAThread] public static void Main() { Application.EnableVisualStyles(); Application.Run(new MacosTodoWatcher()); }
}
