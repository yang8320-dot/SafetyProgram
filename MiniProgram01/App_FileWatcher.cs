/// FILE: MiniProgram01/App_FileWatcher.cs ///
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
    
    // 【新增】用於控制 UI 防洗版的快取，紀錄各資料夾/檔案最後一次顯示在畫面的時間
    private Dictionary<string, DateTime> lastUINotifications = new Dictionary<string, DateTime>();
    
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
        btnGoSet.Click += (s, e) => { new MonitorSettingsWindow(this).Show(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnClear, 1, 0);
        header.Controls.Add(btnGoSet, 2, 0);
        this.Controls.Add(header);

        cardPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        
        cardPanel.Resize += (s, e) => {
            int safeWidth = cardPanel.ClientSize.Width - 25;
            if (safeWidth > 0) {
                foreach (Control c in cardPanel.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
            }
        };

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
        new MonitorListWindow(this).Show();
    }

    private void ReloadAllWatchers() {
        foreach (var w in watchers) { 
            w.EnableRaisingEvents = false; 
            w.Dispose(); 
        }
        watchers.Clear();
        foreach (var val in pathPairs.Values) { 
            StartWatcherFromLine(val); 
        }
    }

    private bool IsSamePath(string p1, string p2) {
        if (string.IsNullOrEmpty(p1) || string.IsNullOrEmpty(p2)) return false;
        return string.Equals(p1.TrimEnd('\\', '/'), p2.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
    }

    private void StartWatcherFromLine(string line) {
        var p = line.Split('|'); 
        if (p.Length < 5) return;
        
        string syncMode = p.Length > 5 ? p[5] : "單向備份";
        
        try {
            if (!Directory.Exists(p[0])) return;
            FileSystemWatcher wSrc = new FileSystemWatcher(p[0]) { 
                IncludeSubdirectories = (p[4] != "僅本層"), 
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size,
                EnableRaisingEvents = true 
            };
            wSrc.Changed += OnFileEvent; 
            wSrc.Created += OnFileEvent; 
            watchers.Add(wSrc);

            if (syncMode == "雙向同步") {
                if (!Directory.Exists(p[1])) Directory.CreateDirectory(p[1]);
                FileSystemWatcher wDst = new FileSystemWatcher(p[1]) { 
                    IncludeSubdirectories = (p[4] != "僅本層"), 
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size,
                    EnableRaisingEvents = true 
                };
                wDst.Changed += OnFileEvent; 
                wDst.Created += OnFileEvent; 
                watchers.Add(wDst);
            }
        } catch { }
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        
        string checkName = Path.GetFileName(e.FullPath);
        
        // 【核心修改 1】：增加排除看圖軟體與系統自動生成的縮圖暫存檔
        if (checkName.StartsWith("~") || 
            checkName.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) ||
            checkName.Equals("Thumbs.db", StringComparison.OrdinalIgnoreCase) ||
            checkName.Equals("desktop.ini", StringComparison.OrdinalIgnoreCase) ||
            checkName.Equals(".DS_Store", StringComparison.OrdinalIgnoreCase)) 
        {
            return;
        }

        string triggerRoot = ((FileSystemWatcher)sender).Path;
        string taskLine = null; 
        bool isTriggerFromSrc = true;

        foreach (var val in pathPairs.Values) {
            var p = val.Split('|');
            if (IsSamePath(triggerRoot, p[0])) { 
                taskLine = val; 
                isTriggerFromSrc = true; 
                break; 
            }
            if (p.Length > 5 && p[5] == "雙向同步" && IsSamePath(triggerRoot, p[1])) { 
                taskLine = val; 
                isTriggerFromSrc = false; 
                break; 
            }
        }
        if (taskLine == null) return; 
        
        var parts = taskLine.Split('|');
        string sourceDir = isTriggerFromSrc ? parts[0] : parts[1];
        string targetDir = isTriggerFromSrc ? parts[1] : parts[0];
        string syncMode = parts.Length > 5 ? parts[5] : "單向備份";

        int targetDepth = ParseDepth(parts[4]);
        if (targetDepth != -1 && GetPathDepth(sourceDir, e.FullPath) > targetDepth) return;

        string cleanSrc = sourceDir.TrimEnd('\\', '/') + "\\";
        string relPath = "";
        try {
            relPath = Uri.UnescapeDataString(new Uri(cleanSrc).MakeRelativeUri(new Uri(e.FullPath)).ToString().Replace('/', '\\'));
        } catch { return; }

        string targetFile = Path.Combine(targetDir, relPath);

        if (syncMode == "雙向同步" && File.Exists(targetFile)) {
            if (Math.Abs((File.GetLastWriteTime(e.FullPath) - File.GetLastWriteTime(targetFile)).TotalSeconds) < 2) return;
        }

        lock (lockObj) {
            // 防抖動處理
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (DateTime.Now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = DateTime.Now;
        }
        
        string customName = parts.Length > 7 ? parts[7] : "";
        DoBackup(e.FullPath, targetFile, parts[2] == "顯示在監控", parts[3], relPath, customName);
    }

    private int GetPathDepth(string root, string file) {
        string rel = file.Replace(root, "").TrimStart('\\', '/');
        return string.IsNullOrEmpty(rel) ? 0 : rel.Split(new char[] { '\\', '/' }).Length - 1;
    }

    private int ParseDepth(string d) {
        if (d == "無限層") return -1; 
        if (d == "僅本層") return 0;
        if (d == "第一層") return 1; 
        if (d == "第二層") return 2; 
        if (d == "第三層") return 3; 
        if (d == "第十層") return 10;
        
        string clean = d.Replace("層", "").Replace("第", "").Trim();
        int res; 
        return int.TryParse(clean, out res) ? res : -1;
    }

    private void DoBackup(string srcFile, string targetFile, bool showUI, string freqStr, string relPath, string customName) {
        ThreadPool.QueueUserWorkItem(_ => {
            int sec; 
            if(!int.TryParse(freqStr.Replace("秒", "").Trim(), out sec)) sec = 1;
            Thread.Sleep(sec * 1000); 

            bool copySuccess = false;
            for(int i = 0; i < 3; i++) { 
                try {
                    string dir = Path.GetDirectoryName(targetFile);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    File.Copy(srcFile, targetFile, true); 
                    copySuccess = true; 
                    break;
                } catch { 
                    Thread.Sleep(800); 
                }
            }

            if (copySuccess && showUI) {
                // 【核心修改 2】：將 UI 顯示改為資料夾收斂模式
                string[] pathParts = relPath.Split(new char[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
                bool isInsideFolder = pathParts.Length > 1;
                
                // 如果在資料夾內，取最上層資料夾名稱；否則直接顯示檔名
                string displayFileName = isInsideFolder ? "📁 " + pathParts[0] + " (資料夾群組)" : Path.GetFileName(srcFile);

                bool shouldUpdateUI = false;
                lock (lockObj) {
                    // 同一個資料夾/檔案，在 2 秒內只會在 UI 產生一次卡片，避免大量檔案同時備份造成洗版
                    if (!lastUINotifications.ContainsKey(displayFileName) || (DateTime.Now - lastUINotifications[displayFileName]).TotalSeconds > 2) {
                        lastUINotifications[displayFileName] = DateTime.Now;
                        shouldUpdateUI = true;
                    }
                }

                if (!shouldUpdateUI) return; // UI 防洗版略過，但背景確實已經備份完了

                try {
                    this.Invoke(new Action(() => {
                        parentForm.AlertTab(0); 
                        
                        Panel c = new Panel() { 
                            Width = 340, AutoSize = true, 
                            MinimumSize = new Size(0, 65), 
                            BackColor = Color.WhiteSmoke, 
                            BorderStyle = BorderStyle.FixedSingle, 
                            Margin = new Padding(0, 2, 0, 5) 
                        };
                        
                        TableLayoutPanel tlp = new TableLayoutPanel() { 
                            Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, 
                            AutoSize = true, Padding = new Padding(2) 
                        };
                        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 55f)); 
                        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 

                        FlowLayoutPanel btnPnl = new FlowLayoutPanel() { 
                            FlowDirection = FlowDirection.TopDown, 
                            AutoSize = true, Margin = new Padding(0) 
                        };
                        
                        Button bView = new Button() { 
                            Text = "查看", Width = 50, Height = 25, 
                            Margin = new Padding(0, 2, 0, 2), 
                            BackColor = AppleBlue, ForeColor = Color.White, 
                            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand 
                        };
                        bView.Click += (s, e2) => {
                            try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + srcFile + "\""); } catch { }
                            cardPanel.Controls.Remove(c); 
                            c.Dispose(); 
                        };
                        
                        Button bClose = new Button() { 
                            Text = "X", Width = 50, Height = 25, 
                            Margin = new Padding(0, 2, 0, 2), 
                            BackColor = Color.IndianRed, ForeColor = Color.White, 
                            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand 
                        };
                        bClose.Click += (s, e2) => { 
                            cardPanel.Controls.Remove(c); 
                            c.Dispose(); 
                        };
                        
                        btnPnl.Controls.Add(bView); 
                        btnPnl.Controls.Add(bClose);

                        // 顯示位置：如果是群組顯示第一層資料夾路徑，若是單一檔案則顯示原路徑
                        string displayLoc = string.IsNullOrWhiteSpace(customName) ? (isInsideFolder ? pathParts[0] : relPath) : customName;
                        Label lbl = new Label() { 
                            Text = displayFileName + "\n位置: " + displayLoc, 
                            Dock = DockStyle.Fill, AutoSize = true, 
                            Padding = new Padding(5, 5, 0, 0), 
                            TextAlign = ContentAlignment.MiddleLeft 
                        };
                        
                        tlp.Controls.Add(btnPnl, 0, 0); 
                        tlp.Controls.Add(lbl, 1, 0); 
                        
                        c.Controls.Add(tlp);
                        cardPanel.Controls.Add(c); 
                        cardPanel.Controls.SetChildIndex(c, 0);
                        
                        if (cardPanel.Controls.Count > 15) { 
                            var old = cardPanel.Controls[15]; 
                            cardPanel.Controls.RemoveAt(15); 
                            old.Dispose(); 
                        }
                    }));
                } catch { } 
            }
        });
    }

    private void RunRetentionSweep() {
        foreach (var val in pathPairs.Values) {
            var p = val.Split('|');
            if (p.Length >= 7 && p[5] == "單向備份" && p[6] != "永久") {
                int months = 0; 
                string mStr = p[6].Replace("個月", "").Replace("月", "").Trim();
                if (int.TryParse(mStr, out months) && months > 0 && Directory.Exists(p[1])) {
                    try {
                        DateTime threshold = DateTime.Now.AddMonths(-months);
                        foreach (string file in Directory.GetFiles(p[1], "*", SearchOption.AllDirectories)) {
                            if (File.GetLastWriteTime(file) < threshold) { 
                                try { File.Delete(file); } catch { } 
                            }
                        }
                    } catch { }
                }
            }
        }
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) return;
        foreach (string line in File.ReadAllLines(configFile)) {
            var parts = line.Split('|'); 
            if (parts.Length >= 5 && Directory.Exists(parts[0])) { 
                pathPairs[parts[0]] = line; 
            }
        }
        ReloadAllWatchers();
    }

    private void SaveAllConfigs() {
        List<string> lines = new List<string>(); 
        foreach (var val in pathPairs.Values) lines.Add(val);
        File.WriteAllLines(configFile, lines);
    }
}

// ==========================================
// 視窗：監控設定 (保持原樣不變)
// ==========================================
public class MonitorSettingsWindow : Form {
    /* 略... (這段完全不需修改，可直接沿用原本的) */
    private App_FileWatcher parentWatcher;
    private TextBox txtCustomName, txtSource, txtBackup;
    private ComboBox cmbMethod, cmbFreq, cmbDepth, cmbSync, cmbRetain;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);
    private string editingKey = null;

    public MonitorSettingsWindow(App_FileWatcher watcher, string existingKey = null, string existingData = null) {
        this.parentWatcher = watcher;
        this.editingKey = existingKey;
        this.Text = string.IsNullOrEmpty(existingKey) ? "新增監控設定" : "編輯監控設定";
        this.Width = 380;
        this.AutoSize = true;
        this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.TopMost = true; 

        FlowLayoutPanel mainFlow = new FlowLayoutPanel() { 
            FlowDirection = FlowDirection.TopDown, 
            WrapContents = false, 
            Padding = new Padding(20), 
            AutoSize = true 
        };
        
        txtCustomName = AddTextRow(mainFlow, "自訂顯示名稱：");
        txtSource = AddPathRow(mainFlow, "來源路徑：");
        txtBackup = AddPathRow(mainFlow, "備份路徑：");
        cmbSync = AddComboRow(mainFlow, "同步模式：", new string[] { "單向備份", "雙向同步" }, false);
        cmbRetain = AddComboRow(mainFlow, "自動刪除(單向)：", new string[] { "永久", "1個月", "3個月", "6個月", "12個月" }, false);
        cmbMethod = AddComboRow(mainFlow, "通知方式：", new string[] { "顯示在監控", "隱藏背後執行" }, false);
        cmbFreq = AddComboRow(mainFlow, "延遲執行：", new string[] { "0秒", "1秒", "3秒", "5秒", "10秒" }, false);
        cmbDepth = AddComboRow(mainFlow, "監控深度：", new string[] { "無限層", "僅本層", "第一層", "第二層", "第三層" }, false);

        Button btnSave = new Button() { 
            Text = string.IsNullOrEmpty(existingKey) ? "新增任務" : "儲存修改", 
            Width = 320, Height = 40, 
            BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, 
            FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 15, 0, 0) 
        };
        
        if (!string.IsNullOrEmpty(existingData)) {
            var p = existingData.Split('|');
            if (p.Length >= 5) {
                txtSource.Text = p[0];
                txtBackup.Text = p[1];
                cmbMethod.Text = p[2];
                cmbFreq.Text = p[3];
                cmbDepth.Text = p[4];
                if (p.Length > 5) cmbSync.Text = p[5];
                if (p.Length > 6) cmbRetain.Text = p[6];
                if (p.Length > 7) txtCustomName.Text = p[7];
            }
        }

        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtSource.Text) || string.IsNullOrWhiteSpace(txtBackup.Text)) {
                MessageBox.Show("來源與備份路徑不可為空！"); 
                return;
            }
            if (string.IsNullOrEmpty(editingKey)) {
                parentWatcher.AddNewTask(txtSource.Text, txtBackup.Text, cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text);
            } else {
                parentWatcher.UpdateTask(editingKey, txtSource.Text, txtBackup.Text, cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text);
            }
            this.Close();
        };

        mainFlow.Controls.Add(btnSave);

        if (string.IsNullOrEmpty(existingKey)) {
            Button btnList = new Button() { 
                Text = "管理現有任務", Width = 320, Height = 40, 
                BackColor = Color.Gray, ForeColor = Color.White, 
                FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 10, 0, 0) 
            };
            btnList.Click += (s, e) => {
                parentWatcher.OpenListWindow();
                this.Close();
            };
            mainFlow.Controls.Add(btnList);
        }
        this.Controls.Add(mainFlow);
    }

    private TextBox AddTextRow(FlowLayoutPanel container, string label) {
        container.Controls.Add(new Label() { Text = label, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        TextBox tb = new TextBox() { Width = 320, Font = MainFont };
        container.Controls.Add(tb);
        return tb;
    }

    private TextBox AddPathRow(FlowLayoutPanel container, string label) {
        container.Controls.Add(new Label() { Text = label, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        FlowLayoutPanel row = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, 10), WrapContents = false };
        TextBox tb = new TextBox() { Width = 230, Font = MainFont };
        
        Button btnSel = new Button() { Text = "選", Width = 38, Height = 28, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnSel.Click += (s, e) => { 
            using(FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; 
            }
        };
        
        Button btnPaste = new Button() { Text = "貼", Width = 38, Height = 28, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnPaste.Click += (s, e) => { 
            string p = Clipboard.GetText().Trim(' ', '\"'); 
            if (!string.IsNullOrEmpty(p)) tb.Text = p; 
        };
        
        row.Controls.Add(tb); 
        row.Controls.Add(btnSel); 
        row.Controls.Add(btnPaste);
        container.Controls.Add(row);
        return tb;
    }

    private ComboBox AddComboRow(FlowLayoutPanel container, string label, string[] items, bool editable) {
        container.Controls.Add(new Label() { Text = label, AutoSize = true, Margin = new Padding(0, 5, 0, 2) });
        ComboBox cb = new ComboBox() { Width = 320, Font = MainFont, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        cb.Items.AddRange(items);
        if (items.Length > 0) cb.SelectedIndex = 0;
        container.Controls.Add(cb);
        return cb;
    }
}

// ==========================================
// 視窗：管理監控任務清單 (保持原樣不變)
// ==========================================
public class MonitorListWindow : Form {
    /* 略... (這段完全不需修改，可直接沿用原本的) */
    private App_FileWatcher parentWatcher;
    private FlowLayoutPanel flow;
    private Font MainFont = new Font("Microsoft JhengHei UI", 9.5f);

    public MonitorListWindow(App_FileWatcher watcher) {
        this.parentWatcher = watcher;
        this.Text = "管理現有任務";
        this.Width = 500;
        this.Height = 500;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.White;
        this.TopMost = true; 

        flow = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, AutoScroll = true, 
            FlowDirection = FlowDirection.TopDown, 
            WrapContents = false, Padding = new Padding(15) 
        };
        
        flow.Resize += (s, e) => {
            int w = flow.ClientSize.Width - 30;
            if (w > 0) foreach (Control c in flow.Controls) if (c is Panel) c.Width = w;
        };
        
        this.Controls.Add(flow);
        RefreshList();
    }

    private void RefreshList() {
        flow.Controls.Clear();
        var pairs = parentWatcher.GetPathPairs();
        
        if (pairs.Count == 0) {
            flow.Controls.Add(new Label() { Text = "目前沒有任何監控任務。", AutoSize = true, Margin = new Padding(10), Font = MainFont });
            return;
        }

        foreach (var kvp in pairs) {
            string key = kvp.Key;
            string data = kvp.Value;
            var parts = data.Split('|');
            if (parts.Length < 5) continue;

            string customName = parts.Length > 7 ? parts[7] : "";
            string displayName = string.IsNullOrWhiteSpace(customName) ? parts[0] : string.Format("{0}\n({1})", customName, parts[0]);

            Panel card = new Panel() { 
                AutoSize = true, MinimumSize = new Size(0, 60), 
                BackColor = Color.FromArgb(248, 248, 250), 
                Margin = new Padding(0, 0, 0, 10), Padding = new Padding(10) 
            };
            card.Width = flow.ClientSize.Width > 30 ? flow.ClientSize.Width - 30 : 400;

            TableLayoutPanel tlp = new TableLayoutPanel() { 
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, AutoSize = true 
            };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110f));

            Label lbl = new Label() { Text = displayName + "\n=> " + parts[1], Dock = DockStyle.Fill, AutoSize = true, Font = MainFont };
            
            FlowLayoutPanel btnPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0) };
            
            Button btnEdit = new Button() { Text = "編輯", Width = 50, Height = 30, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnEdit.Click += (s, e) => { new MonitorSettingsWindow(parentWatcher, key, data).Show(); this.Close(); };
            
            Button btnDel = new Button() { Text = "刪除", Width = 50, Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnDel.Click += (s, e) => {
                if (MessageBox.Show("確定要刪除此監控任務嗎？", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    parentWatcher.DeleteTask(key); RefreshList();
                }
            };
            
            btnPanel.Controls.Add(btnEdit); btnPanel.Controls.Add(btnDel);
            tlp.Controls.Add(lbl, 0, 0); tlp.Controls.Add(btnPanel, 1, 0);
            card.Controls.Add(tlp); flow.Controls.Add(card);
        }
    }
}
