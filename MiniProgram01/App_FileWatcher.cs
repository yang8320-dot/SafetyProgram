/*
 * 檔案功能：檔案監控與自動備份/同步模組 (支援單向備份、雙向同步、即時監控)
 * 對應選單名稱：檔案監控
 * 對應資料庫名稱：(本模組採用純文字檔存儲) MainDB_FileWatcher.txt
 * 資料表名稱：無 (資料欄位採用 '|' 符號間隔)
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_FileWatcher : UserControl
{
    private MainForm parentForm;
    private ContextMenu trayMenu;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MainDB_FileWatcher.txt");

    // --- 核心資料結構 ---
    // key: SourcePath, value: 設定字串 (src|dst|method|freq|depth|syncMode|retention|customName)
    private Dictionary<string, string> pathPairs = new Dictionary<string, string>();
    private Dictionary<string, FileSystemWatcher> activeWatchers = new Dictionary<string, FileSystemWatcher>();

    // --- 介面控制項 ---
    private FlowLayoutPanel taskPanel;

    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleGreen = Color.FromArgb(52, 199, 89);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);
    private static Font SmallFont = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Regular);

    public App_FileWatcher(MainForm parent, ContextMenu trayMenu)
    {
        this.parentForm = parent;
        this.trayMenu = trayMenu;

        // 1. 初始化控制項與 DPI 支援
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15);

        // 2. 建構純程式碼 UI
        InitializeUI();

        // 3. 載入資料並啟動監控
        _ = LoadAndStartWatchersAsync();
    }

    /// <summary>
    /// 建構 iOS 風格純程式碼介面 (Code-First UI)
    /// </summary>
    private void InitializeUI()
    {
        // ==========================================
        // 頂部標題與控制區塊
        // ==========================================
        TableLayoutPanel header = new TableLayoutPanel()
        {
            Dock = DockStyle.Top,
            Height = 45,
            ColumnCount = 2,
            BackColor = Color.Transparent
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label()
        {
            Text = "檔案監控與同步",
            Font = new Font(MainFont.FontFamily, 14f, FontStyle.Bold),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(5, 0, 0, 0)
        };

        Button btnAdd = new Button()
        {
            Text = "新增任務",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = AppleBlue,
            ForeColor = Color.White,
            Font = BoldFont,
            Margin = new Padding(0, 5, 0, 5)
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += (s, e) => { new MonitorSettingsWindow(this, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);

        // ==========================================
        // 中間列表區塊 (卡片容器)
        // ==========================================
        taskPanel = new FlowLayoutPanel()
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = AppleBgColor
        };

        taskPanel.Resize += (s, e) =>
        {
            int safeWidth = taskPanel.ClientSize.Width - 20;
            if (safeWidth > 0)
            {
                foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
            }
        };

        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent });
        this.Controls.Add(header);
        
        taskPanel.BringToFront();
    }

    // ==========================================
    // 核心邏輯：存檔與載入 (Thread-Safety)
    // ==========================================
    private async Task LoadAndStartWatchersAsync()
    {
        if (File.Exists(configFile))
        {
            try
            {
                string[] lines = await Task.Run(() => File.ReadAllLines(configFile));
                pathPairs.Clear();
                foreach (var line in lines)
                {
                    string[] parts = line.Split('|');
                    if (parts.Length >= 8)
                    {
                        pathPairs[parts[0]] = line;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入監控設定失敗: {ex.Message}");
            }
        }
        
        ReloadAllWatchers();
        RefreshUI();
    }

    public async Task SaveAllConfigsAsync()
    {
        try
        {
            List<string> lines = pathPairs.Values.ToList();
            await Task.Run(() => File.WriteAllLines(configFile, lines));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"儲存監控設定失敗: {ex.Message}");
        }
    }

    public Dictionary<string, string> GetPathPairs() => pathPairs;

    // ==========================================
    // 核心邏輯：背景監控引擎
    // ==========================================
    public void ReloadAllWatchers()
    {
        // 停止並清除舊有監控
        foreach (var watcher in activeWatchers.Values)
        {
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
        }
        activeWatchers.Clear();

        // 重新啟動所有合法任務
        foreach (var kvp in pathPairs)
        {
            string[] parts = kvp.Value.Split('|');
            string src = parts[0];
            string syncMode = parts[5];

            if (Directory.Exists(src))
            {
                try
                {
                    FileSystemWatcher fsw = new FileSystemWatcher(src)
                    {
                        IncludeSubdirectories = parts[4] == "包含子目錄",
                        NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.DirectoryName,
                        EnableRaisingEvents = true
                    };

                    fsw.Created += (s, e) => OnFileChanged(e, kvp.Value);
                    fsw.Changed += (s, e) => OnFileChanged(e, kvp.Value);
                    fsw.Renamed += (s, e) => OnFileChanged(e, kvp.Value);
                    // 雙向同步或需鏡像刪除時可於此擴充 Deleted 事件

                    activeWatchers[src] = fsw;
                }
                catch { /* 忽略無法監控的目錄 */ }
            }
        }
    }

    private void OnFileChanged(FileSystemEventArgs e, string configData)
    {
        // 在背景執行緒中處理，確保不會卡住主程式
        Task.Run(() => 
        {
            string[] parts = configData.Split('|');
            string srcDir = parts[0];
            string targetDir = parts[1];
            string syncMode = parts[5];

            if (!Directory.Exists(targetDir)) return;

            try
            {
                string cleanSrc = srcDir.TrimEnd('\\', '/') + "\\";
                string relPath = Uri.UnescapeDataString(new Uri(cleanSrc).MakeRelativeUri(new Uri(e.FullPath)).ToString().Replace('/', '\\'));
                string targetFile = Path.Combine(targetDir, relPath);

                // 若為雙向同步，檢查時間避免無限迴圈觸發
                if (syncMode == "雙向同步" && File.Exists(targetFile))
                {
                    if (Math.Abs((File.GetLastWriteTime(e.FullPath) - File.GetLastWriteTime(targetFile)).TotalSeconds) < 2)
                        return; 
                }

                Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                File.Copy(e.FullPath, targetFile, true);
                
                // 觸發主視窗閃爍警報 (呼叫時必須透過 Invoke 確保 Thread-Safety)
                if (parentForm != null && parentForm.InvokeRequired == false)
                {
                    // 此處假設 FileWatcher 在 TabControl 中的 Index，實務上可透過事件委派傳遞
                    parentForm.Invoke(new Action(() => parentForm.AddAlertTab(0))); 
                }
            }
            catch { /* 檔案可能被鎖定，等待下一次觸發 */ }
        });
    }

    // ==========================================
    // CRUD 任務管理介面更新
    // ==========================================
    public async Task AddOrUpdateTaskAsync(string oldSrc, string newSrc, string dst, string method, string freq, string depth, string syncMode, string retention, string customName)
    {
        if (!string.IsNullOrEmpty(oldSrc) && oldSrc != newSrc && pathPairs.ContainsKey(newSrc))
        {
            MessageBox.Show("新的來源路徑已存在監控清單中！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!string.IsNullOrEmpty(oldSrc)) pathPairs.Remove(oldSrc);
        
        pathPairs[newSrc] = $"{newSrc}|{dst}|{method}|{freq}|{depth}|{syncMode}|{retention}|{customName}";
        
        await SaveAllConfigsAsync();
        ReloadAllWatchers();
        RefreshUI();
    }

    public void RefreshUI()
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(RefreshUI));
            return;
        }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var kvp in pathPairs)
        {
            var parts = kvp.Value.Split('|');
            if (parts.Length < 8) continue;

            string src = parts[0];
            string dst = parts[1];
            string syncMode = parts[5];
            string customName = parts[7];
            bool isRunning = activeWatchers.ContainsKey(src);

            // iOS 卡片設計
            Panel card = new Panel()
            {
                Width = startWidth,
                AutoSize = true,
                MinimumSize = new Size(0, 85),
                Margin = new Padding(0, 0, 0, 15),
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            TableLayoutPanel tlp = new TableLayoutPanel()
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                AutoSize = true
            };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f)); // 按鈕區寬度

            // 標題與狀態
            Label lblTitle = new Label()
            {
                Text = $"{(string.IsNullOrWhiteSpace(customName) ? "未命名任務" : customName)}  [{syncMode}]",
                Font = BoldFont,
                AutoSize = true,
                ForeColor = isRunning ? AppleGreen : Color.Gray,
                Dock = DockStyle.Top,
                Margin = new Padding(0, 0, 0, 5)
            };
            tlp.Controls.Add(lblTitle, 0, 0);

            // 來源與目標路徑資訊
            Label lblDetails = new Label()
            {
                Text = $"來源: {src}\n目標: {dst}",
                Font = SmallFont,
                ForeColor = Color.DarkGray,
                AutoSize = true,
                Dock = DockStyle.Top
            };
            tlp.SetRowSpan(lblDetails, 2);
            tlp.Controls.Add(lblDetails, 0, 1);

            // 編輯按鈕
            Button btnEdit = new Button()
            {
                Text = "編輯",
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = AppleBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = SmallFont,
                Margin = new Padding(10, 0, 0, 5)
            };
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (s, e) => new MonitorSettingsWindow(this, kvp.Value).ShowDialog();
            tlp.Controls.Add(btnEdit, 1, 0);

            // 刪除按鈕
            Button btnDel = new Button()
            {
                Text = "刪除",
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = AppleRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = SmallFont,
                Margin = new Padding(10, 0, 0, 0)
            };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += async (s, e) =>
            {
                if (MessageBox.Show("確定要刪除這個監控任務嗎？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    pathPairs.Remove(src);
                    await SaveAllConfigsAsync();
                    ReloadAllWatchers();
                    RefreshUI();
                }
            };
            tlp.Controls.Add(btnDel, 1, 1);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }
}

// ==========================================
// 視窗：新增/編輯監控設定 (iOS 風格純程式碼 UI)
// ==========================================
public class MonitorSettingsWindow : Form
{
    private App_FileWatcher parent;
    private string originalSource = null;

    private TextBox txtCustomName, txtSource, txtBackup;
    private ComboBox cmbMethod, cmbFreq, cmbDepth, cmbSync, cmbRetain;

    public MonitorSettingsWindow(App_FileWatcher parent, string configData)
    {
        this.parent = parent;
        
        // 視窗基本設定
        this.Text = configData == null ? "新增監控任務" : "編輯監控任務";
        this.Width = 500;
        this.Height = 620;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247); // Apple BgColor
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        // 主容器
        FlowLayoutPanel flow = new FlowLayoutPanel()
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            Padding = new Padding(20),
            WrapContents = false,
            AutoScroll = true
        };

        // 解析舊有資料
        string cName = "", src = "", dst = "", method = "即時監控", freq = "即時", depth = "包含子目錄", syncMode = "單向備份", retain = "永久保留";
        if (configData != null)
        {
            var p = configData.Split('|');
            if (p.Length >= 8)
            {
                originalSource = p[0]; src = p[0]; dst = p[1]; method = p[2]; freq = p[3];
                depth = p[4]; syncMode = p[5]; retain = p[6]; cName = p[7];
            }
        }

        // --- 表單欄位建立方法 ---
        void AddLabel(string text)
        {
            flow.Controls.Add(new Label() { Text = text, AutoSize = true, Margin = new Padding(0, 10, 0, 5) });
        }

        TextBox AddTextBox(string val)
        {
            TextBox tb = new TextBox() { Text = val, Width = 440, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(5) };
            flow.Controls.Add(tb);
            return tb;
        }

        ComboBox AddComboBox(string[] items, string sel)
        {
            ComboBox cb = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 440 };
            cb.Items.AddRange(items);
            if (cb.Items.Contains(sel)) cb.SelectedItem = sel; else cb.SelectedIndex = 0;
            flow.Controls.Add(cb);
            return cb;
        }

        // --- 建立 UI 欄位 ---
        AddLabel("任務自訂名稱：");
        txtCustomName = AddTextBox(cName);

        AddLabel("來源目錄 (需監控的資料夾)：");
        txtSource = AddTextBox(src);
        txtSource.DoubleClick += (s, e) => {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK) txtSource.Text = fbd.SelectedPath;
        };
        flow.Controls.Add(new Label() { Text = "(雙擊輸入框可開啟資料夾選擇視窗)", ForeColor = Color.Gray, Font = new Font(this.Font.FontFamily, 8.5f) });

        AddLabel("目標目錄 (備份/同步至)：");
        txtBackup = AddTextBox(dst);
        txtBackup.DoubleClick += (s, e) => {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK) txtBackup.Text = fbd.SelectedPath;
        };

        AddLabel("同步模式：");
        cmbSync = AddComboBox(new string[] { "單向備份", "雙向同步" }, syncMode);

        AddLabel("監控深度：");
        cmbDepth = AddComboBox(new string[] { "僅限當前目錄", "包含子目錄" }, depth);

        // --- 隱藏較少使用的設定，設為預設 ---
        cmbMethod = new ComboBox() { Text = method };
        cmbFreq = new ComboBox() { Text = freq };
        cmbRetain = new ComboBox() { Text = retain };

        // 儲存按鈕
        Button btnSave = new Button()
        {
            Text = "儲存設定",
            Width = 440,
            Height = 45,
            BackColor = Color.FromArgb(0, 122, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold),
            Margin = new Padding(0, 25, 0, 0)
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtSource.Text) || string.IsNullOrWhiteSpace(txtBackup.Text))
            {
                MessageBox.Show("來源與目標目錄不可為空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            await parent.AddOrUpdateTaskAsync(
                originalSource, txtSource.Text.Trim(), txtBackup.Text.Trim(), 
                cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, 
                cmbSync.Text, cmbRetain.Text, txtCustomName.Text.Trim()
            );

            this.Close();
        };

        flow.Controls.Add(btnSave);
        this.Controls.Add(flow);
    }
}
