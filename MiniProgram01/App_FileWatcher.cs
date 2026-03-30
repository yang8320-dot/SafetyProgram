using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private ContextMenu trayMenu;
    private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
    private FlowLayoutPanel cardPanel;
    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    
    // 儲存監控路徑與對應備份路徑的字典
    private Dictionary<string, string> pathPairs = new Dictionary<string, string>();

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9f);

    public App_FileWatcher(MainForm mainForm, ContextMenu menu) {
        this.parentForm = mainForm;
        this.trayMenu = menu;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // --- UI 配置 ---
        Label lblHint = new Label() { 
            Text = "📢 檔案變動將依原目錄結構自動備份", 
            Dock = DockStyle.Top, Height = 30, Font = new Font(MainFont, FontStyle.Bold), ForeColor = Color.Gray 
        };
        this.Controls.Add(lblHint);

        cardPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White 
        };
        this.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        LoadConfigAndStartWatch();
    }

    private void LoadConfigAndStartWatch() {
        if (!File.Exists(configFile)) {
            // 建立範例設定檔
            File.WriteAllLines(configFile, new string[] { "C:\\SourcePath|C:\\BackupPath" });
            return;
        }

        foreach (string line in File.ReadAllLines(configFile)) {
            string[] parts = line.Split('|');
            if (parts.Length >= 2) {
                string src = parts[0].Trim();
                string dst = parts[1].Trim();
                if (Directory.Exists(src)) {
                    pathPairs[src] = dst;
                    StartWatcher(src);
                }
            }
        }
    }

    private void StartWatcher(string path) {
        FileSystemWatcher watcher = new FileSystemWatcher(path);
        watcher.IncludeSubdirectories = true; // 預設監控子資料夾
        watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
        watcher.Changed += OnFileChanged;
        watcher.Created += OnFileChanged;
        watcher.Renamed += OnFileChanged;
        watcher.EnableRaisingEvents = true;
        watchers.Add(watcher);
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return; // 略過資料夾變動，只處理檔案

        string srcRoot = ((FileSystemWatcher)sender).Path;
        if (pathPairs.ContainsKey(srcRoot)) {
            string dstRoot = pathPairs[srcRoot];
            DoBackupWithStructure(srcRoot, e.FullPath, dstRoot);
        }
    }

    // ==========================================
    // 【核心修正】依實際資料夾結構進行備份
    // ==========================================
    private void DoBackupWithStructure(string srcRoot, string fileFullPath, string dstRoot) {
        try {
            // 1. 計算相對路徑 (例如: SubDir\test.txt)
            // 使用 Uri 類別來安全地計算路徑差異，避免手動切字串出錯
            Uri rootUri = new Uri(srcRoot.EndsWith("\\") ? srcRoot : srcRoot + "\\");
            Uri fileUri = new Uri(fileFullPath);
            string relativePath = Uri.UnescapeDataString(rootUri.MakeRelativeUri(fileUri).ToString().Replace('/', '\\'));

            // 2. 組合備份目標路徑
            string destFile = Path.Combine(dstRoot, relativePath);
            string destDir = Path.GetDirectoryName(destFile);

            // 3. 自動建立遺失的子資料夾
            if (!Directory.Exists(destDir)) {
                Directory.CreateDirectory(destDir);
            }

            // 4. 執行備份 (覆蓋模式)
            // 稍微延遲一下，避免檔案被其他程式鎖定中 (例如剛存檔完)
            System.Threading.Thread.Sleep(500); 
            File.Copy(fileFullPath, destFile, true);

            // 5. 在介面顯示卡片
            this.Invoke(new Action(() => AddCard(Path.GetFileName(fileFullPath), relativePath, destFile)));
        }
        catch (Exception ex) {
            Console.WriteLine("備份出錯: " + ex.Message);
        }
    }

    private void AddCard(string fileName, string relPath, string fullDest) {
        Panel card = new Panel() { 
            Width = 320, Height = 60, BackColor = Color.FromArgb(250, 250, 252), 
            Margin = new Padding(0, 5, 0, 5), BorderStyle = BorderStyle.FixedSingle 
        };
        
        Label lblInfo = new Label() { 
            Text = string.Format("{0}\n路徑: {1}", fileName, relPath), 
            Location = new Point(10, 10), Width = 230, Height = 40, Font = MainFont 
        };
        
        Button btnOpen = new Button() { 
            Text = "查看", Location = new Point(250, 15), Width = 55, Height = 30, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White 
        };
        btnOpen.Click += (s, e) => {
            try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + fullDest + "\""); } catch { }
        };

        card.Controls.Add(lblInfo);
        card.Controls.Add(btnOpen);
        
        // 新卡片插在最上面
        cardPanel.Controls.Add(card);
        cardPanel.Controls.SetChildIndex(card, 0);

        // 如果卡片太多，自動移除舊的 (保持 20 個就好)
        if (cardPanel.Controls.Count > 20) cardPanel.Controls.RemoveAt(20);
    }

    // 輔助函式：讓其他模組共用
    public static string ShowInputBox(string prompt, string title, string defaultValue) {
        Form form = new Form() { Width = 380, Height = 170, Text = title, StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog };
        Label lbl = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true, Font = MainFont };
        TextBox txt = new TextBox() { Left = 20, Top = 50, Width = 320, Text = defaultValue, Font = MainFont };
        Button btnOk = new Button() { Text = "確定", Left = 250, Top = 90, Width = 90, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
        form.Controls.AddRange(new Control[] { lbl, txt, btnOk });
        return form.ShowDialog() == DialogResult.OK ? txt.Text : "";
    }
}
