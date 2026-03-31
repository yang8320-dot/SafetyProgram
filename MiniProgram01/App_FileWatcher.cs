using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel cardPanel;

    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private Dictionary<string, string> pathPairs = new Dictionary<string, string>();
    private Dictionary<string, DateTime> lastProcessedTimes = new Dictionary<string, DateTime>();
    private readonly object lockObj = new object();
    private System.Windows.Forms.Timer retentionTimer; 

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(5);

        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 3 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f)); 
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label() { Text = "異動紀錄清單", Font = new Font(MainFont, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(10,0,0,0) };
        Button btnClear = new Button() { Text = "一鍵清除", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Margin = new Padding(2,8,2,8), Cursor = Cursors.Hand };
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();
        
        Button btnGoSet = new Button() { Text = "設定參數", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, BackColor = Color.Gainsboro, Margin = new Padding(2,8,8,8), Cursor = Cursors.Hand };
        btnGoSet.Click += (s, e) => { new MonitorSettingsWindow(this).ShowDialog(); };

        header.Controls.AddRange(new Control[] { lblTitle, btnClear, btnGoSet });
        this.Controls.Add(header);

        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        this.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        LoadConfigAndStartWatch();

        retentionTimer = new System.Windows.Forms.Timer() { Interval = 3600000, Enabled = true };
        retentionTimer.Tick += (s, e) => RunRetentionSweep();
        RunRetentionSweep(); 
    }

    public Dictionary<string, string> GetPathPairs() { return pathPairs; }

    public void AddNewTask(string src, string dst, string method, string freq, string depth, string syncMode, string retention, string customName) {
        pathPairs[src] = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}", src, dst, method, freq, depth, syncMode, retention, customName);
        SaveAllConfigs(); 
        ReloadAllWatchers(); 
        MessageBox.Show("監控任務已成功建立！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void UpdateTask(string oldSrc, string newSrc, string dst, string method, string freq, string depth, string syncMode, string retention, string customName) {
        if (!IsSamePath(oldSrc, newSrc) && pathPairs.ContainsKey(newSrc)) {
            MessageBox.Show("新的來源路徑已存在監控清單中！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        pathPairs.Remove(oldSrc);
        pathPairs[newSrc] = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}", newSrc, dst, method, freq, depth, syncMode, retention, customName);
        SaveAllConfigs();
        ReloadAllWatchers(); 
        MessageBox.Show("任務修改成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void DeleteTask(string key) {
        pathPairs.Remove(key);
        SaveAllConfigs();
        ReloadAllWatchers(); 
    }

    public void OpenListWindow() {
        using(var listForm = new MonitorListWindow(this)) { listForm.ShowDialog(); }
    }

    private void ReloadAllWatchers() {
        foreach (var w in watchers) { w.EnableRaisingEvents = false; w.Dispose(); }
        watchers.Clear();
        foreach (var val in pathPairs.Values) { StartWatcherFromLine(val); }
    }

    private bool IsSamePath(string p1, string p2) {
        if (string.IsNullOrEmpty(p1) || string.IsNullOrEmpty(p2)) return false;
        return string.Equals(p1.TrimEnd('\\', '/'), p2.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
    }

    private void StartWatcherFromLine(string line) {
        var p = line.Split('|'); if (p.Length < 5) return;
        string syncMode = p.Length > 5 ? p[5] : "單向備份";
        
        FileSystemWatcher wSrc = new FileSystemWatcher(p[0]) { 
            IncludeSubdirectories = (p[4] != "僅本層"), 
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size,
            EnableRaisingEvents = true 
        };
        wSrc.Changed += OnFileEvent; wSrc.Created += OnFileEvent; watchers.Add(wSrc);

        if (syncMode == "雙向同步") {
            if (!Directory.Exists(p[1])) Directory.CreateDirectory(p[1]);
            FileSystemWatcher wDst = new FileSystemWatcher(p[1]) { 
                IncludeSubdirectories = (p[4] != "僅本層"), 
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size,
                EnableRaisingEvents = true 
            };
            wDst.Changed += OnFileEvent; wDst.Created += OnFileEvent; watchers.Add(wDst);
        }
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        
        string checkName = Path.GetFileName(e.FullPath);
        if (checkName.StartsWith("~$") || checkName.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase)) return;

        string triggerRoot = ((FileSystemWatcher)sender).Path;
        string taskLine = null; bool isTriggerFromSrc = true;

        foreach (var val in pathPairs.Values) {
            var p = val.Split('|');
            if (IsSamePath(triggerRoot, p[0])) { taskLine = val; isTriggerFromSrc = true; break; }
            if (p.Length > 5 && p[5] == "雙向同步" && IsSamePath(triggerRoot, p[1])) { taskLine = val; isTriggerFromSrc = false; break; }
        }
        if (taskLine == null) return; 
        
        var parts = taskLine.Split('|');
        string sourceDir = isTriggerFromSrc ? parts[0] : parts[1];
        string targetDir = isTriggerFromSrc ? parts[1] : parts[0];
        string syncMode = parts.Length > 5 ? parts[5] : "單向備份";

        int targetDepth = ParseDepth(parts[4]);
        if (targetDepth != -1 && GetPathDepth(sourceDir, e.FullPath) > targetDepth) return;

        string cleanSrc = sourceDir.TrimEnd('\\', '/') + "\\";
        string rel = Uri.UnescapeDataString(new Uri(cleanSrc).MakeRelativeUri(new Uri(e.FullPath)).ToString().Replace('/', '\\'));
        string targetFile = Path.Combine(targetDir, rel);

        if (syncMode == "雙向同步" && File.Exists(targetFile)) {
            if (Math.Abs((File.GetLastWriteTime(e.FullPath) - File.GetLastWriteTime(targetFile)).TotalSeconds) < 2) return;
        }

        lock (lockObj) {
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (DateTime.Now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = DateTime.Now;
        }
        
        string customName = parts.Length > 7 ? parts[7] : "";
        DoBackup(e.FullPath, targetFile, parts[2] == "顯示在監控", parts[3], rel, customName);
    }

    private int GetPathDepth(string root, string file) {
        string rel = file.Replace(root, "").TrimStart('\\', '/');
        return string.IsNullOrEmpty(rel) ? 0 : rel.Split(new char[] { '\\', '/' }).Length - 1;
    }

    private int ParseDepth(string d) {
        if (d == "無限層") return -1; if (d == "僅本層") return 0;
        if (d == "第一層") return 1; if (d == "第二層") return 2; if (d == "第三層") return 3; if (d == "第十層") return 10;
        string clean = d.Replace("層", "").Replace("第", "").Trim();
        int res; return int.TryParse(clean, out res) ? res : -1;
    }

    // 【修改核心】：將按鈕移至左側，上下兩行顯示
    private void DoBackup(string srcFile, string targetFile, bool showUI, string freqStr, string relPath, string customName) {
        ThreadPool.QueueUserWorkItem(_ => {
            int sec; if(!int.TryParse(freqStr.Replace("秒", "").Trim(), out sec)) sec = 1;
            Thread.Sleep(sec * 1000); 

            bool copySuccess = false;
            for(int i = 0; i < 3; i++) { 
                try {
                    if (!Directory.Exists(Path.GetDirectoryName(targetFile))) Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                    File.Copy(srcFile, targetFile, true); 
                    copySuccess = true; break;
                } catch { Thread.Sleep(800); }
            }

            if (copySuccess && showUI) {
                try {
                    this.Invoke(new Action(() => {
                        parentForm.AlertTab(0); 
                        
                        // 外層卡片容器
                        Panel c = new Panel() { Width = 340, AutoSize = true, MinimumSize = new Size(340, 65), BackColor = Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 2, 0, 5) };
                        
                        // 劃分左右兩塊：左邊按鈕區，右邊文字區
                        TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, AutoSize = true, Padding = new Padding(2) };
                        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 55f)); // 左側保留 55px 給按鈕
                        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // 右側全部給文字

                        // 左側按鈕區 (上下排列)
                        FlowLayoutPanel btnPnl = new FlowLayoutPanel() { FlowDirection = FlowDirection.TopDown, AutoSize = true, Margin = new Padding(0) };
                        
                        Button bView = new Button() { Text = "查看", Width = 50, Height = 25, Margin = new Padding(0, 2, 0, 2), BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                        bView.Click += (s, e2) => {
                            System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + srcFile + "\"");
                            cardPanel.Controls.Remove(c); c.Dispose(); 
                        };
                        
                        Button bClose = new Button() { Text = "X", Width = 50, Height = 25, Margin = new Padding(0, 2, 0, 2), BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                        bClose.Click += (s, e2) => { cardPanel.Controls.Remove(c); c.Dispose(); };
                        
                        btnPnl.Controls.Add(bView); 
                        btnPnl.Controls.Add(bClose);

                        // 右側文字區
                        string displayLoc = string.IsNullOrWhiteSpace(customName) ? relPath : customName;
                        Label lbl = new Label() { Text = Path.GetFileName(srcFile) + "\n位置: " + displayLoc, Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(5, 5, 0, 0), TextAlign = ContentAlignment.MiddleLeft };
                        
                        // 將左右兩塊放入 TableLayoutPanel
                        tlp.Controls.Add(btnPnl, 0, 0); 
                        tlp.Controls.Add(lbl, 1, 0); 
                        
                        c.Controls.Add(tlp);
                        cardPanel.Controls.Add(c); cardPanel.Controls.SetChildIndex(c, 0);
                        
                        if (cardPanel.Controls.Count > 15) { var old = cardPanel.Controls[15]; cardPanel.Controls.RemoveAt(15); old.Dispose(); }
                    }));
                } catch { } 
            }
        });
    }

    private void RunRetentionSweep() {
        foreach (var val in pathPairs.Values) {
            var p = val.Split('|');
            if (p.Length >= 7 && p[5] == "單向備份" && p[6] != "永久") {
                int months = 0; string mStr = p[6].Replace("個月", "").Replace("月", "").Trim();
                if (int.TryParse(mStr, out months) && months > 0 && Directory.Exists(p[1])) {
                    try {
                        DateTime threshold = DateTime.Now.AddMonths(-months);
                        foreach (string file in Directory.GetFiles(p[1], "*", SearchOption.AllDirectories)) {
                            if (File.GetLastWriteTime(file) < threshold) { try { File.Delete(file); } catch { } }
                        }
                    } catch { }
                }
            }
        }
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            var parts = line.Split('|'); if (parts.Length >= 5 && Directory.Exists(parts[0])) { pathPairs[parts[0]] = line; }
        }
        ReloadAllWatchers();
    }

    private void SaveAllConfigs() {
        List<string> lines = new List<string>(); foreach (var val in pathPairs.Values) lines.Add(val);
        File.WriteAllLines(configFile, lines);
    }
}

public class MonitorSettingsWindow : Form {
    private App_FileWatcher parentWatcher;
    private TextBox txtCustomName, txtSource, txtBackup; 
    private ComboBox cmbMethod, cmbFreq, cmbDepth, cmbSync, cmbRetain;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public MonitorSettingsWindow(App_FileWatcher watcher) {
        this.parentWatcher = watcher; 
        this.Text = "監控設定"; 
        this.Width = 380; this.AutoSize = true; this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        this.StartPosition = FormStartPosition.CenterScreen; this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;

        FlowLayoutPanel mainFlow = new FlowLayoutPanel() { FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20), AutoSize = true };

        GroupBox gbNew = new GroupBox() { Text = "新增監控路徑", Font = new Font(MainFont, FontStyle.Bold), Width = 320, AutoSize = true, Margin = new Padding(0,0,0,15) };
        FlowLayoutPanel flowNew = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5, 15, 5, 10), AutoSize = true, Font = MainFont };
        
        txtCustomName = AddTextRow(flowNew, "自訂顯示名稱："); 
        txtSource = AddPathRow(flowNew, "來源路徑："); 
        txtBackup = AddPathRow(flowNew, "備份路徑：");
        cmbSync = AddComboRow(flowNew, "同步模式：", new string[] { "單向備份", "雙向同步" }, false);
        cmbRetain = AddComboRow(flowNew, "自動刪除(單向)：", new string[] { "永久", "1個月", "3個月", "6個月" }, true);
        cmbMethod = AddComboRow(flowNew, "提示方式：", new string[] { "顯示在監控", "背景執行" }, false);
        cmbFreq = AddComboRow(flowNew, "頻率(秒)：", new string[] { "1", "3", "5", "10", "30", "60" }, true);
        cmbDepth = AddComboRow(flowNew, "監控深度：", new string[] { "無限層", "僅本層", "第一層", "第二層", "第三層", "第十層" }, true);

        Button btnAdd = new Button() { Text = "+ 新增至監控任務", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(5, 10, 0, 5), Cursor = Cursors.Hand };
        btnAdd.Click += (s, e) => {
            if (!Directory.Exists(txtSource.Text.Trim())) { MessageBox.Show("來源路徑無效！"); return; }
            parentWatcher.AddNewTask(txtSource.Text.Trim(), txtBackup.Text.Trim(), cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text.Trim());
            txtSource.Text = ""; txtBackup.Text = ""; txtCustomName.Text = "";
        };
        flowNew.Controls.Add(btnAdd); gbNew.Controls.Add(flowNew); mainFlow.Controls.Add(gbNew);

        GroupBox gbManage = new GroupBox() { Text = "管理中心", Font = new Font(MainFont, FontStyle.Bold), Width = 320, AutoSize = true };
        FlowLayoutPanel flowManage = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5, 15, 5, 10), AutoSize = true, Font = MainFont };
        Button btnShowList = new Button() { Text = "開啟監控項目清冊", Width = 290, Height = 45, FlatStyle = FlatStyle.Flat, BackColor = Color.WhiteSmoke, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(5,0,0,0), Cursor = Cursors.Hand }; 
        btnShowList.Click += (s, e) => parentWatcher.OpenListWindow();
        flowManage.Controls.Add(btnShowList); gbManage.Controls.Add(flowManage); mainFlow.Controls.Add(gbManage);
        this.Controls.Add(mainFlow);
    }

    private TextBox AddTextRow(FlowLayoutPanel container, string labelText) {
        container.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        TextBox tb = new TextBox() { Width = 255 }; row.Controls.Add(tb); container.Controls.Add(row); return tb;
    }

    private TextBox AddPathRow(FlowLayoutPanel container, string labelText) {
        container.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        TextBox tb = new TextBox() { Width = 170 }; 
        Button btnSel = new Button() { Text = "選", Width = 35, Height = 25, Font = new Font(MainFont.FontFamily, 8f) };
        btnSel.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; };
        Button btnPaste = new Button() { Text = "貼", Width = 35, Height = 25, Font = new Font(MainFont.FontFamily, 8f) };
        btnPaste.Click += (s, e) => { string p = Clipboard.GetText().Trim(' ', '\"'); if (!string.IsNullOrEmpty(p)) tb.Text = p; };
        row.Controls.AddRange(new Control[] { tb, btnSel, btnPaste });
        container.Controls.Add(row); return tb;
    }

    private ComboBox AddComboRow(FlowLayoutPanel container, string labelText, string[] items, bool editable) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 5, 0, 5) };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        ComboBox cb = new ComboBox() { Width = 150, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        cb.Items.AddRange(items); cb.SelectedIndex = 0;
        row.Controls.Add(cb); container.Controls.Add(row); return cb;
    }
}

public class MonitorListWindow : Form {
    private App_FileWatcher parentWatcher;
    private FlowLayoutPanel list;

    public MonitorListWindow(App_FileWatcher watcher) {
        this.parentWatcher = watcher;
        this.Text = "監控項目管理中心"; 
        this.Width = 800; this.Height = 850; this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White; this.Font = new Font("Microsoft JhengHei UI", 10f);

        list = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        this.Controls.Add(list);
        ReloadUI();
    }

    public void ReloadUI() {
        list.Controls.Clear();
        var data = parentWatcher.GetPathPairs();
        
        foreach (var key in data.Keys) {
            var p = data[key].Split('|');
            
            Panel card = new Panel() { Width = 740, AutoSize = true, MinimumSize = new Size(740, 0), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 10), BackColor = Color.FromArgb(248, 248, 250) };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, AutoSize = true, Padding = new Padding(10) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180f));

            FlowLayoutPanel txtFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            txtFlow.Controls.Add(new Label() { Text = "源：" + p[0], AutoSize = true, ForeColor = Color.FromArgb(0, 102, 204) });
            txtFlow.Controls.Add(new Label() { Text = "備：" + p[1], AutoSize = true, ForeColor = Color.FromArgb(0, 153, 76) });
            
            string syncMode = p.Length > 5 ? p[5] : "單向備份";
            string retention = p.Length > 6 ? p[6] : "永久";
            string customName = p.Length > 7 ? p[7] : ""; 
            string nameDisplay = string.IsNullOrWhiteSpace(customName) ? "未設定" : customName;

            txtFlow.Controls.Add(new Label() { Text = string.Format("名稱：{0} | 同步：{1} | 提醒：{2} | 刪除：{3}", nameDisplay, syncMode, p[2], retention), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 8.5f), ForeColor = Color.DimGray, Margin = new Padding(0, 5, 0, 0) });
            
            FlowLayoutPanel btnFlow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 0) };
            
            Button btnEdit = new Button() { Text = "調整", Width = 75, Height = 35, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnEdit.Click += (s, e) => { new EditMonitorTaskWindow(parentWatcher, this, key, data[key]).ShowDialog(); };
            
            Button btnDel = new Button() { Text = "移除", Width = 75, Height = 35, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnDel.Click += (s, e) => { if(MessageBox.Show("確定要永久移除此監控任務？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { parentWatcher.DeleteTask(key); ReloadUI(); } };
            
            btnFlow.Controls.Add(btnEdit); btnFlow.Controls.Add(btnDel);
            tlp.Controls.Add(txtFlow, 0, 0); tlp.Controls.Add(btnFlow, 1, 0);
            card.Controls.Add(tlp); list.Controls.Add(card);
        }
    }
}

public class EditMonitorTaskWindow : Form {
    private App_FileWatcher parentWatcher;
    private MonitorListWindow listWindow;
    private string originalSrcKey;

    private TextBox txtCustomName, txtSource, txtBackup; 
    private ComboBox cmbMethod, cmbFreq, cmbDepth, cmbSync, cmbRetain;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public EditMonitorTaskWindow(App_FileWatcher watcher, MonitorListWindow lw, string key, string taskData) {
        this.parentWatcher = watcher; this.listWindow = lw; this.originalSrcKey = key;
        this.Text = "修改監控任務"; 
        this.Width = 380; this.AutoSize = true; this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        this.StartPosition = FormStartPosition.CenterParent; this.BackColor = Color.White; this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;

        FlowLayoutPanel mainFlow = new FlowLayoutPanel() { FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(20), AutoSize = true };
        GroupBox gbEdit = new GroupBox() { Text = "編輯參數", Font = new Font(MainFont, FontStyle.Bold), Width = 320, AutoSize = true };
        FlowLayoutPanel flowEdit = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5, 15, 5, 10), AutoSize = true, Font = MainFont };

        txtCustomName = AddTextRow(flowEdit, "自訂顯示名稱："); 
        txtSource = AddPathRow(flowEdit, "來源路徑："); 
        txtBackup = AddPathRow(flowEdit, "備份路徑：");
        cmbSync = AddComboRow(flowEdit, "同步模式：", new string[] { "單向備份", "雙向同步" }, false);
        cmbRetain = AddComboRow(flowEdit, "自動刪除(單向)：", new string[] { "永久", "1個月", "3個月", "6個月" }, true);
        cmbMethod = AddComboRow(flowEdit, "提示方式：", new string[] { "顯示在監控", "背景執行" }, false);
        cmbFreq = AddComboRow(flowEdit, "頻率(秒)：", new string[] { "1", "3", "5", "10", "30", "60" }, true);
        cmbDepth = AddComboRow(flowEdit, "監控深度：", new string[] { "無限層", "僅本層", "第一層", "第二層", "第三層", "第十層" }, true);

        var p = taskData.Split('|');
        txtSource.Text = p[0]; txtBackup.Text = p[1]; cmbMethod.Text = p[2]; cmbFreq.Text = p[3]; cmbDepth.Text = p[4];
        if (p.Length > 5) cmbSync.Text = p[5]; 
        if (p.Length > 6) cmbRetain.Text = p[6];
        if (p.Length > 7) txtCustomName.Text = p[7]; 

        Button btnSave = new Button() { Text = "儲存修改", Width = 290, Height = 40, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Margin = new Padding(5, 10, 0, 5), Cursor = Cursors.Hand }; 
        btnSave.Click += (s, e) => {
            if (!Directory.Exists(txtSource.Text.Trim())) { MessageBox.Show("來源路徑無效！"); return; }
            parentWatcher.UpdateTask(originalSrcKey, txtSource.Text.Trim(), txtBackup.Text.Trim(), cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text.Trim());
            listWindow.ReloadUI(); 
            this.Close();
        };
        
        flowEdit.Controls.Add(btnSave); gbEdit.Controls.Add(flowEdit); mainFlow.Controls.Add(gbEdit);
        this.Controls.Add(mainFlow);
    }

    private TextBox AddTextRow(FlowLayoutPanel container, string labelText) {
        container.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        TextBox tb = new TextBox() { Width = 255 }; row.Controls.Add(tb); container.Controls.Add(row); return tb;
    }

    private TextBox AddPathRow(FlowLayoutPanel container, string labelText) {
        container.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
        TextBox tb = new TextBox() { Width = 170 }; 
        Button btnSel = new Button() { Text = "選", Width = 35, Height = 25, Font = new Font(MainFont.FontFamily, 8f) };
        btnSel.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; };
        Button btnPaste = new Button() { Text = "貼", Width = 35, Height = 25, Font = new Font(MainFont.FontFamily, 8f) };
        btnPaste.Click += (s, e) => { string p = Clipboard.GetText().Trim(' ', '\"'); if (!string.IsNullOrEmpty(p)) tb.Text = p; };
        row.Controls.AddRange(new Control[] { tb, btnSel, btnPaste });
        container.Controls.Add(row); return tb;
    }

    private ComboBox AddComboRow(FlowLayoutPanel container, string labelText, string[] items, bool editable) {
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 5, 0, 5) };
        row.Controls.Add(new Label() { Text = labelText, AutoSize = true, Margin = new Padding(0, 7, 0, 0) });
        ComboBox cb = new ComboBox() { Width = 150, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        cb.Items.AddRange(items); cb.SelectedIndex = 0;
        row.Controls.Add(cb); container.Controls.Add(row); return cb;
    }
}
