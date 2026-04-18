// ============================================================
// FILE: MiniProgram01/App_Shortcuts.cs 
// ============================================================
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

public class App_Shortcuts : UserControl {
    private MainForm parentForm;
    private string shortcutFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_shortcuts.txt");
    private FlowLayoutPanel taskPanel;

    // 拖曳排序相關變數
    private int dragInsertIndex = -1;
    
    // --- iOS 風格色彩與字體定義 ---
    private static Color iosBackground = Color.FromArgb(242, 242, 247);
    private static Color iosCardWhite = Color.White;
    private static Color iosAppleBlue = Color.FromArgb(0, 122, 255);
    private static Color iosGreen = Color.FromArgb(52, 199, 89);
    private static Color iosRed = Color.FromArgb(255, 59, 48);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    public class ShortcutItem {
        public string Name { get; set; }
        public string Path { get; set; }
    }
    public List<ShortcutItem> shortcuts = new List<ShortcutItem>();

    public App_Shortcuts(MainForm mainForm) {
        this.parentForm = mainForm;
        this.BackColor = iosBackground; 
        this.Padding = new Padding(15); 
        this.AutoScaleMode = AutoScaleMode.Dpi; // 核心：支援高 DPI 縮放

        TableLayoutPanel header = new TableLayoutPanel() { 
            Dock = DockStyle.Top, 
            Height = 50, 
            ColumnCount = 2 
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        Label lblTitle = new Label() { 
            Text = "常用捷徑", 
            Font = new Font("Microsoft JhengHei UI", 14f, FontStyle.Bold), 
            ForeColor = Color.FromArgb(28, 28, 30),
            Dock = DockStyle.Fill, 
            TextAlign = ContentAlignment.MiddleLeft, 
            Padding = new Padding(5, 0, 0, 0) 
        };
        
        Button btnAdd = new Button() { 
            Text = "新增", 
            Dock = DockStyle.Fill, 
            FlatStyle = FlatStyle.Flat, 
            Margin = new Padding(2, 8, 2, 8), 
            Cursor = Cursors.Hand, 
            BackColor = iosAppleBlue, 
            ForeColor = Color.White,
            Font = BoldFont
        };
        btnAdd.FlatAppearance.BorderSize = 0; 
        btnAdd.Click += (s, e) => { new EditShortcutWindow(this, -1, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);
        this.Controls.Add(header);

        // 初始化容器並開啟拖放功能
        taskPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            AutoScroll = true, 
            FlowDirection = FlowDirection.TopDown, 
            WrapContents = false, 
            BackColor = iosBackground, 
            AllowDrop = true,
            Padding = new Padding(0, 10, 0, 0)
        };
        
        // 綁定拖曳相關事件
        taskPanel.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskPanel.DragOver += OnTaskDragOver;
        taskPanel.DragLeave += (s, e) => { dragInsertIndex = -1; taskPanel.Invalidate(); };
        taskPanel.DragDrop += OnTaskDragDrop;
        taskPanel.Paint += OnTaskContainerPaint;
        
        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - 25;
            if (safeWidth > 0) {
                foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
            }
        };

        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadShortcuts();
    }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 25 : 450;
        
        foreach (var s in shortcuts) {
            Panel card = new Panel() { 
                Width = startWidth, 
                AutoSize = true, 
                MinimumSize = new Size(0, 55), 
                Margin = new Padding(5, 5, 5, 10), 
                BackColor = iosCardWhite, 
                BorderStyle = BorderStyle.None 
            };
            
            TableLayoutPanel tlp = new TableLayoutPanel() { 
                Dock = DockStyle.Fill, 
                ColumnCount = 4, 
                RowCount = 1, 
                AutoSize = true, 
                Padding = new Padding(10) 
            };
            
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45f)); 

            Button btnOpen = new Button() { 
                Text = "開啟", 
                Dock = DockStyle.Fill, 
                Height = 35, 
                BackColor = iosGreen, 
                ForeColor = Color.White, 
                FlatStyle = FlatStyle.Flat, 
                Cursor = Cursors.Hand, 
                Margin = new Padding(0, 0, 5, 0), 
                Font = BoldFont 
            };
            btnOpen.FlatAppearance.BorderSize = 0; 
            btnOpen.Click += (sender, e) => {
                try { 
                    ProcessStartInfo psi = new ProcessStartInfo() { FileName = s.Path, UseShellExecute = true };
                    Process.Start(psi); 
                } 
                catch { MessageBox.Show("無法開啟此捷徑，請檢查路徑或檔案是否存在！", "開啟失敗", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };

            Button btnDel = new Button() { 
                Text = "✕", 
                Dock = DockStyle.Fill, 
                Height = 35, 
                BackColor = iosRed, 
                ForeColor = Color.White, 
                FlatStyle = FlatStyle.Flat, 
                Cursor = Cursors.Hand, 
                Margin = new Padding(0, 0, 5, 0), 
                Font = BoldFont 
            };
            btnDel.FlatAppearance.BorderSize = 0; 
            btnDel.Click += (sender, e) => { 
                if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK) {
                    shortcuts.Remove(s);
                    SaveShortcuts();
                    RefreshUI();
                }
            };

            // 標題文字：新增滑鼠按下的事件來啟動拖曳
            Label lbl = new Label() { 
                Text = s.Name, 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleLeft, 
                AutoSize = true, 
                Font = MainFont, 
                Padding = new Padding(10, 5, 0, 5),
                Cursor = Cursors.SizeAll 
            };
            lbl.MouseDown += (sender, e) => {
                if (e.Button == MouseButtons.Left) card.DoDragDrop(card, DragDropEffects.Move);
            };

            Button btnEdit = new Button() { 
                Text = "修", 
                Dock = DockStyle.Fill, 
                Height = 35, 
                BackColor = iosAppleBlue, 
                ForeColor = Color.White, 
                FlatStyle = FlatStyle.Flat, 
                Cursor = Cursors.Hand, 
                Margin = new Padding(5, 0, 0, 0), 
                Font = BoldFont 
            };
            btnEdit.FlatAppearance.BorderSize = 0; 
            btnEdit.Click += (sender, e) => {
                int idx = shortcuts.IndexOf(s);
                new EditShortcutWindow(this, idx, s).ShowDialog();
            };

            tlp.Controls.Add(btnOpen, 0, 0);
            tlp.Controls.Add(btnDel, 1, 0);
            tlp.Controls.Add(lbl, 2, 0);
            tlp.Controls.Add(btnEdit, 3, 0);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    // --- 拖曳排序核心邏輯 ---

    private void OnTaskDragOver(object sender, DragEventArgs e) {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskPanel.PointToClient(new Point(e.X, e.Y));
        Control target = taskPanel.GetChildAtPoint(clientPoint);
        
        if (target != null) {
            int idx = taskPanel.Controls.GetChildIndex(target);
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            if (dragInsertIndex != idx) { dragInsertIndex = idx; taskPanel.Invalidate(); }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e) {
        if (dragInsertIndex != -1 && taskPanel.Controls.Count > 0) {
            int y = (dragInsertIndex < taskPanel.Controls.Count) ?
                taskPanel.Controls[dragInsertIndex].Top - 2 : taskPanel.Controls[taskPanel.Controls.Count - 1].Bottom + 2;
            e.Graphics.FillRectangle(new SolidBrush(iosAppleBlue), 5, y, taskPanel.Width - 30, 4); 
        }
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        Panel draggedCard = (Panel)e.Data.GetData(typeof(Panel));
        if (draggedCard != null && dragInsertIndex != -1) {
            int targetIdx = dragInsertIndex;
            int currentIdx = taskPanel.Controls.GetChildIndex(draggedCard);
            if (currentIdx < targetIdx) targetIdx--; 
            
            var item = shortcuts[currentIdx];
            shortcuts.RemoveAt(currentIdx);
            
            int finalIdx = Math.Min(targetIdx, shortcuts.Count);
            shortcuts.Insert(finalIdx, item);

            dragInsertIndex = -1; 
            taskPanel.Invalidate(); 
            SaveShortcuts(); 
            RefreshUI(); 
        }
    }

    // --- 存檔與載入 ---

    public void SaveShortcuts() {
        List<string> lines = new List<string>();
        foreach(var s in shortcuts) {
            lines.Add(string.Format("{0}|{1}", s.Name, Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(s.Path))));
        }
        File.WriteAllLines(shortcutFile, lines);
    }

    public void LoadShortcuts() {
        if(!File.Exists(shortcutFile)) return;
        shortcuts.Clear();
        foreach(var l in File.ReadAllLines(shortcutFile)) {
            var p = l.Split('|');
            if(p.Length >= 2) {
                try {
                    string path = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(p[1]));
                    shortcuts.Add(new ShortcutItem() { Name = p[0], Path = path });
                } catch {
                    shortcuts.Add(new ShortcutItem() { Name = p[0], Path = p[1] });
                }
            }
        }
        RefreshUI();
    }
}

// ==========================================
// 視窗：新增/編輯捷徑
// ==========================================
public class EditShortcutWindow : Form {
    private App_Shortcuts parent;
    private int index;
    private TextBox txtName, txtPath;

    public EditShortcutWindow(App_Shortcuts p, int idx, App_Shortcuts.ShortcutItem item) {
        this.parent = p;
        this.index = idx;
        this.Text = idx == -1 ? "新增捷徑" : "編輯捷徑";
        this.Width = 450; 
        this.Height = 280;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false; 
        this.MinimizeBox = false;
        this.BackColor = Color.White; 
        this.AutoScaleMode = AutoScaleMode.Dpi; 
        this.Font = new Font("Microsoft JhengHei UI", 10.5f);

        FlowLayoutPanel f = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            FlowDirection = FlowDirection.TopDown, 
            Padding = new Padding(25) 
        };
        
        f.Controls.Add(new Label() { Text = "捷徑名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5), ForeColor = Color.FromArgb(142, 142, 147) });
        txtName = new TextBox() { 
            Width = 380, 
            Text = item?.Name ?? "", 
            Margin = new Padding(0, 0, 0, 20),
            BorderStyle = BorderStyle.FixedSingle
        };
        f.Controls.Add(txtName);
        
        f.Controls.Add(new Label() { Text = "目標路徑 (檔案 / 資料夾 / 網址)：", AutoSize = true, Margin = new Padding(0, 0, 0, 5), ForeColor = Color.FromArgb(142, 142, 147) });
        
        TableLayoutPanel pathRow = new TableLayoutPanel() { 
            Width = 380, 
            Height = 38, 
            ColumnCount = 2, 
            Margin = new Padding(0, 0, 0, 20) 
        };
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));
        
        txtPath = new TextBox() { 
            Dock = DockStyle.Fill, 
            Text = item?.Path ?? "",
            BorderStyle = BorderStyle.FixedSingle
        };
        
        Button btnBrowse = new Button() { 
            Text = "瀏覽", 
            Dock = DockStyle.Fill, 
            FlatStyle = FlatStyle.Flat, 
            Cursor = Cursors.Hand,
            BackColor = Color.FromArgb(242, 242, 247), 
            FlatAppearance = { BorderSize = 0 }
        };
        btnBrowse.Click += (s, e) => {
            OpenFileDialog ofd = new OpenFileDialog() { Title = "選擇捷徑目標檔案" };
            if (ofd.ShowDialog() == DialogResult.OK) { txtPath.Text = ofd.FileName; }
        };
        
        pathRow.Controls.Add(txtPath, 0, 0);
        pathRow.Controls.Add(btnBrowse, 1, 0);
        f.Controls.Add(pathRow);

        Button btnSave = new Button() { 
            Text = "儲存設定", 
            Width = 380, 
            Height = 42, 
            BackColor = Color.FromArgb(0, 122, 255), 
            ForeColor = Color.White, 
            FlatStyle = FlatStyle.Flat, 
            Cursor = Cursors.Hand,
            Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold)
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtPath.Text)) {
                MessageBox.Show("名稱與路徑不可為空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (index == -1) {
                parent.shortcuts.Add(new App_Shortcuts.ShortcutItem() { Name = txtName.Text, Path = txtPath.Text });
            } else {
                parent.shortcuts[index].Name = txtName.Text;
                parent.shortcuts[index].Path = txtPath.Text;
            }
            parent.SaveShortcuts();
            parent.RefreshUI();
            this.Close();
        };
        f.Controls.Add(btnSave);
        this.Controls.Add(f);
    }
}
