/*
 * 檔案功能：檔案監控與自動備份/同步模組 (SQLite 升級版)
 * 對應選單名稱：檔案監控
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表名稱：FileWatcher
 */

using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_FileWatcher : UserControl
{
    private MainForm parentForm;
    private ContextMenu trayMenu;

    // --- 核心資料結構 ---
    // key: SourcePath, value: 設定陣列 (src, dst, method, freq, depth, syncMode, retention, customName)
    private Dictionary<string, string[]> pathPairs = new Dictionary<string, string[]>();
    private Dictionary<string, FileSystemWatcher> activeWatchers = new Dictionary<string, FileSystemWatcher>();

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

        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15);

        InitializeUI();
        _ = LoadAndStartWatchersAsync();
    }

    private void InitializeUI()
    {
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 2, BackColor = Color.Transparent };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label() { Text = "檔案監控與同步", Font = new Font(MainFont.FontFamily, 14f, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5, 0, 0, 0) };
        Button btnAdd = new Button() { Text = "新增任務", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = AppleBlue, ForeColor = Color.White, Font = BoldFont, Margin = new Padding(0, 5, 0, 5) };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += (s, e) => { new MonitorSettingsWindow(this, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = AppleBgColor };
        taskPanel.Resize += (s, e) =>
        {
            int safeWidth = taskPanel.ClientSize.Width - 20;
            if (safeWidth > 0) { foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth; }
        };

        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent });
        this.Controls.Add(header);
        taskPanel.BringToFront();
    }

    // ==========================================
    // SQLite 資料存取 (Async & Thread-Safety)
    // ==========================================
    private async Task LoadAndStartWatchersAsync()
    {
        try
        {
            var loadedPairs = await Task.Run(() =>
            {
                var dict = new Dictionary<string, string[]>();
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT SourcePath, TargetPath, SyncMethod, Frequency, Depth, SyncMode, Retention, CustomName FROM FileWatcher", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string src = reader.GetString(0);
                            dict[src] = new string[] { 
                                src, reader.GetString(1), reader.GetString(2), reader.GetString(3), 
                                reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetString(7) 
                            };
                        }
                    }
                }
                return dict;
            });

            pathPairs = loadedPairs;
            ReloadAllWatchers();
            RefreshUI();
        }
        catch (Exception ex) { MessageBox.Show($"載入監控設定失敗: {ex.Message}"); }
    }

    public async Task AddOrUpdateTaskAsync(string oldSrc, string[] data)
    {
        string newSrc = data[0];
        if (!string.IsNullOrEmpty(oldSrc) && oldSrc != newSrc && pathPairs.ContainsKey(newSrc))
        {
            MessageBox.Show("新的來源路徑已存在監控清單中！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
        }

        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        // 若修改了來源路徑，先刪除舊的
                        if (!string.IsNullOrEmpty(oldSrc) && oldSrc != newSrc)
                        {
                            using (var cmdDel = new SQLiteCommand("DELETE FROM FileWatcher WHERE SourcePath = @OldSrc", conn, transaction))
                            {
                                cmdDel.Parameters.AddWithValue("@OldSrc", oldSrc);
                                cmdDel.ExecuteNonQuery();
                            }
                        }

                        // 新增或更新 (INSERT OR REPLACE)
                        using (var cmdIns = new SQLiteCommand(@"INSERT OR REPLACE INTO FileWatcher 
                            (SourcePath, TargetPath, SyncMethod, Frequency, Depth, SyncMode, Retention, CustomName) 
                            VALUES (@Src, @Dst, @Method, @Freq, @Depth, @SyncMode, @Retain, @CName)", conn, transaction))
                        {
                            cmdIns.Parameters.AddWithValue("@Src", data[0]); cmdIns.Parameters.AddWithValue("@Dst", data[1]);
                            cmdIns.Parameters.AddWithValue("@Method", data[2]); cmdIns.Parameters.AddWithValue("@Freq", data[3]);
                            cmdIns.Parameters.AddWithValue("@Depth", data[4]); cmdIns.Parameters.AddWithValue("@SyncMode", data[5]);
                            cmdIns.Parameters.AddWithValue("@Retain", data[6]); cmdIns.Parameters.AddWithValue("@CName", data[7]);
                            cmdIns.ExecuteNonQuery();
                        }
                        transaction.Commit();
                    }
                }
            });

            if (!string.IsNullOrEmpty(oldSrc)) pathPairs.Remove(oldSrc);
            pathPairs[newSrc] = data;

            ReloadAllWatchers();
            RefreshUI();
        }
        catch (Exception ex) { MessageBox.Show($"儲存監控設定失敗: {ex.Message}"); }
    }

    public async Task DeleteTaskAsync(string srcPath)
    {
        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("DELETE FROM FileWatcher WHERE SourcePath = @Src", conn))
                    {
                        cmd.Parameters.AddWithValue("@Src", srcPath);
                        cmd.ExecuteNonQuery();
                    }
                }
            });

            pathPairs.Remove(srcPath);
            ReloadAllWatchers();
            RefreshUI();
        }
        catch (Exception ex) { MessageBox.Show($"刪除失敗: {ex.Message}"); }
    }

    // ==========================================
    // 核心邏輯：背景監控引擎
    // ==========================================
    public void ReloadAllWatchers()
    {
        foreach (var watcher in activeWatchers.Values) { watcher.EnableRaisingEvents = false; watcher.Dispose(); }
        activeWatchers.Clear();

        foreach (var kvp in pathPairs)
        {
            string[] parts = kvp.Value;
            string src = parts[0];

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

                    fsw.Created += (s, e) => OnFileChanged(e, parts);
                    fsw.Changed += (s, e) => OnFileChanged(e, parts);
                    fsw.Renamed += (s, e) => OnFileChanged(e, parts);

                    activeWatchers[src] = fsw;
                }
                catch { /* 忽略無法監控的目錄 */ }
            }
        }
    }

    private void OnFileChanged(FileSystemEventArgs e, string[] parts)
    {
        Task.Run(() => 
        {
            string srcDir = parts[0]; string targetDir = parts[1]; string syncMode = parts[5];
            if (!Directory.Exists(targetDir)) return;

            try
            {
                string cleanSrc = srcDir.TrimEnd('\\', '/') + "\\";
                string relPath = Uri.UnescapeDataString(new Uri(cleanSrc).MakeRelativeUri(new Uri(e.FullPath)).ToString().Replace('/', '\\'));
                string targetFile = Path.Combine(targetDir, relPath);

                if (syncMode == "雙向同步" && File.Exists(targetFile))
                {
                    if (Math.Abs((File.GetLastWriteTime(e.FullPath) - File.GetLastWriteTime(targetFile)).TotalSeconds) < 2) return; 
                }

                Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                File.Copy(e.FullPath, targetFile, true);
                
                if (parentForm != null) { parentForm.AddAlertTab(0); }
            }
            catch { }
        });
    }

    // ==========================================
    // UI 更新
    // ==========================================
    public void RefreshUI()
    {
        if (this.InvokeRequired) { this.Invoke(new Action(RefreshUI)); return; }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var kvp in pathPairs)
        {
            var parts = kvp.Value;
            string src = parts[0]; string dst = parts[1]; string syncMode = parts[5]; string customName = parts[7];
            bool isRunning = activeWatchers.ContainsKey(src);

            Panel card = new Panel() { Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 85), Margin = new Padding(0, 0, 0, 15), BackColor = Color.White, Padding = new Padding(10) };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3, AutoSize = true };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f)); 

            Label lblTitle = new Label() { Text = $"{(string.IsNullOrWhiteSpace(customName) ? "未命名任務" : customName)}  [{syncMode}]", Font = BoldFont, AutoSize = true, ForeColor = isRunning ? AppleGreen : Color.Gray, Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 5) };
            tlp.Controls.Add(lblTitle, 0, 0);

            Label lblDetails = new Label() { Text = $"來源: {src}\n目標: {dst}", Font = SmallFont, ForeColor = Color.DarkGray, AutoSize = true, Dock = DockStyle.Top };
            tlp.SetRowSpan(lblDetails, 2); tlp.Controls.Add(lblDetails, 0, 1);

            Button btnEdit = new Button() { Text = "編輯", Dock = DockStyle.Top, Height = 30, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(10, 0, 0, 5) };
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (s, e) => new MonitorSettingsWindow(this, parts).ShowDialog();
            tlp.Controls.Add(btnEdit, 1, 0);

            Button btnDel = new Button() { Text = "刪除", Dock = DockStyle.Top, Height = 30, BackColor = AppleRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(10, 0, 0, 0) };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += async (s, e) => { if (MessageBox.Show("確定要刪除這個任務嗎？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { await DeleteTaskAsync(src); } };
            tlp.Controls.Add(btnDel, 1, 1);

            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }
}

// ==========================================
// 視窗：新增/編輯監控設定
// ==========================================
public class MonitorSettingsWindow : Form
{
    private App_FileWatcher parent;
    private string originalSource = null;

    private TextBox txtCustomName, txtSource, txtBackup;
    private ComboBox cmbMethod, cmbFreq, cmbDepth, cmbSync, cmbRetain;

    public MonitorSettingsWindow(App_FileWatcher parent, string[] configData)
    {
        this.parent = parent;
        
        this.Text = configData == null ? "新增監控任務" : "編輯監控任務";
        this.Width = 500; this.Height = 620; this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247); this.AutoScaleMode = AutoScaleMode.Dpi; this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20), WrapContents = false, AutoScroll = true };

        string cName = "", src = "", dst = "", method = "即時監控", freq = "即時", depth = "包含子目錄", syncMode = "單向備份", retain = "永久保留";
        if (configData != null)
        {
            originalSource = configData[0]; src = configData[0]; dst = configData[1]; method = configData[2]; freq = configData[3];
            depth = configData[4]; syncMode = configData[5]; retain = configData[6]; cName = configData[7];
        }

        void AddLabel(string text) { flow.Controls.Add(new Label() { Text = text, AutoSize = true, Margin = new Padding(0, 10, 0, 5) }); }
        TextBox AddTextBox(string val) { TextBox tb = new TextBox() { Text = val, Width = 440, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(5) }; flow.Controls.Add(tb); return tb; }
        ComboBox AddComboBox(string[] items, string sel) { ComboBox cb = new ComboBox() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 440 }; cb.Items.AddRange(items); cb.SelectedItem = cb.Items.Contains(sel) ? sel : items[0]; flow.Controls.Add(cb); return cb; }

        AddLabel("任務自訂名稱："); txtCustomName = AddTextBox(cName);
        AddLabel("來源目錄 (需監控的資料夾)："); txtSource = AddTextBox(src);
        txtSource.DoubleClick += (s, e) => { FolderBrowserDialog fbd = new FolderBrowserDialog(); if (fbd.ShowDialog() == DialogResult.OK) txtSource.Text = fbd.SelectedPath; };
        flow.Controls.Add(new Label() { Text = "(雙擊輸入框可開啟資料夾選擇視窗)", ForeColor = Color.Gray, Font = new Font(this.Font.FontFamily, 8.5f) });

        AddLabel("目標目錄 (備份/同步至)："); txtBackup = AddTextBox(dst);
        txtBackup.DoubleClick += (s, e) => { FolderBrowserDialog fbd = new FolderBrowserDialog(); if (fbd.ShowDialog() == DialogResult.OK) txtBackup.Text = fbd.SelectedPath; };

        AddLabel("同步模式："); cmbSync = AddComboBox(new string[] { "單向備份", "雙向同步" }, syncMode);
        AddLabel("監控深度："); cmbDepth = AddComboBox(new string[] { "僅限當前目錄", "包含子目錄" }, depth);

        cmbMethod = new ComboBox() { Text = method }; cmbFreq = new ComboBox() { Text = freq }; cmbRetain = new ComboBox() { Text = retain };

        Button btnSave = new Button() { Text = "儲存設定", Width = 440, Height = 45, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), Margin = new Padding(0, 25, 0, 0) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtSource.Text) || string.IsNullOrWhiteSpace(txtBackup.Text)) { MessageBox.Show("來源與目標不可為空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            string[] newData = new string[] { txtSource.Text.Trim(), txtBackup.Text.Trim(), cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text.Trim() };
            await parent.AddOrUpdateTaskAsync(originalSource, newData);
            this.Close();
        };

        flow.Controls.Add(btnSave); this.Controls.Add(flow);
    }
}
