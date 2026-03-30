using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

public class App_Shortcuts : UserControl {
    private MainForm parentForm;
    private string shortcutsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "shortcuts.txt");
    private FlowLayoutPanel listPanel;
    private TextBox txtPath;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 9f);

    private class ShortcutItem {
        public string DisplayName;
        public string TargetPath;
    }
    private List<ShortcutItem> items = new List<ShortcutItem>();

    public App_Shortcuts(MainForm mainForm) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // --- 頂部輸入區 ---
        FlowLayoutPanel topPanel = new FlowLayoutPanel() { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };

        Label lblTitle = new Label() { Text = "新增捷徑 (路徑或拖放)：", AutoSize = true, Font = MainFont, Margin = new Padding(0, 0, 0, 5) };
        
        txtPath = new TextBox() { Width = 320, Font = MainFont, BorderStyle = BorderStyle.FixedSingle };
        
        FlowLayoutPanel btnRow = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 5, 0, 10) };
        Button btnFile = new Button() { Text = "選檔案", Width = 75, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray };
        Button btnFolder = new Button() { Text = "選資料夾", Width = 85, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.LightGray };
        Button btnAdd = new Button() { Text = "確認新增", Width = 150, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) };

        btnFile.Click += (s, e) => { using (OpenFileDialog ofd = new OpenFileDialog()) { if (ofd.ShowDialog() == DialogResult.OK) txtPath.Text = ofd.FileName; } };
        btnFolder.Click += (s, e) => { using (FolderBrowserDialog fbd = new FolderBrowserDialog()) { if (fbd.ShowDialog() == DialogResult.OK) txtPath.Text = fbd.SelectedPath; } };
        btnAdd.Click += (s, e) => { AddShortcut(txtPath.Text); };

        btnRow.Controls.AddRange(new Control[] { btnFile, btnFolder, btnAdd });
        topPanel.Controls.AddRange(new Control[] { lblTitle, txtPath, btnRow });
        this.Controls.Add(topPanel);

        // --- 捷徑顯示區 ---
        listPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        this.Controls.Add(listPanel);
        listPanel.BringToFront();

        LoadShortcuts();
    }

    private void AddShortcut(string path) {
        if (string.IsNullOrWhiteSpace(path)) return;
        string name = Path.GetFileName(path);
        if (string.IsNullOrEmpty(name)) name = path; // 處理磁碟根目錄

        items.Add(new ShortcutItem() { DisplayName = name, TargetPath = path });
        SaveShortcuts();
        RefreshUI();
        txtPath.Text = "";
    }

    private void RefreshUI() {
        listPanel.Controls.Clear();
        foreach (var item in items) {
            Panel card = new Panel() { Width = 330, Height = 40, Margin = new Padding(0, 2, 0, 2), BackColor = Color.FromArgb(250, 250, 252) };
            
            // 顯示名稱 (點擊可編輯)
            Label lblName = new Label() { Text = item.DisplayName, Location = new Point(5, 10), Width = 160, AutoEllipsis = true, Font = MainFont, Cursor = Cursors.Hand };
            lblName.Click += (s, e) => {
                string newName = App_FileWatcher.ShowInputBox("修改顯示名稱：", "✏️ 重新命名", item.DisplayName);
                if (!string.IsNullOrEmpty(newName)) { item.DisplayName = newName; SaveShortcuts(); RefreshUI(); }
            };

            // 開啟按鈕
            Button btnOpen = new Button() { Text = "開啟", Location = new Point(170, 6), Width = 70, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
            btnOpen.Click += (s, e) => { try { Process.Start(new ProcessStartInfo(item.TargetPath) { UseShellExecute = true }); } catch { MessageBox.Show("找不到路徑或檔案！"); } };

            // 刪除按鈕
            Button btnDel = new Button() { Text = "✕", Location = new Point(245, 6), Width = 30, Height = 28, FlatStyle = FlatStyle.Flat, ForeColor = Color.Red };
            btnDel.Click += (s, e) => { items.Remove(item); SaveShortcuts(); RefreshUI(); };

            card.Controls.AddRange(new Control[] { lblName, btnOpen, btnDel });
            listPanel.Controls.Add(card);
        }
    }

    private void LoadShortcuts() {
        if (!File.Exists(shortcutsFile)) return;
        foreach (string line in File.ReadAllLines(shortcutsFile)) {
            string[] parts = line.Split('|');
            if (parts.Length >= 2) items.Add(new ShortcutItem() { DisplayName = parts[0], TargetPath = parts[1] });
        }
        RefreshUI();
    }

    private void SaveShortcuts() {
        List<string> lines = new List<string>();
        foreach (var item in items) lines.Add(item.DisplayName + "|" + item.TargetPath);
        File.WriteAllLines(shortcutsFile, lines.ToArray());
    }
}
