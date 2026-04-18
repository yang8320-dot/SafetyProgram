// ============================================================
// FILE: MiniProgram01/App_TodoList.cs 
// ============================================================
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Text;

public class App_TodoList : UserControl {
    private MainForm parentForm;
    private string listType;
    private string transferLabel;
    private string dataFile;
    
    // 用於雙向轉移的目標列表 (待辦 <-> 待規)
    public App_TodoList TargetList { get; set; }
    
    private FlowLayoutPanel taskPanel;
    private int dragInsertIndex = -1;

    // --- iOS 風格色彩與字體定義 ---
    private static Color iosBackground = Color.FromArgb(242, 242, 247);
    private static Color iosCardWhite = Color.White;
    private static Color iosAppleBlue = Color.FromArgb(0, 122, 255);
    private static Color iosRed = Color.FromArgb(255, 59, 48);
    private static Color iosGrayText = Color.FromArgb(142, 142, 147);
    private static Color iosYellow = Color.FromArgb(255, 204, 0);
    private static Color iosGreen = Color.FromArgb(52, 199, 89);
    
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    // --- 資料模型 ---
    public class TodoTask {
        public string Name { get; set; }
        public string ColorStr { get; set; }
        public string Tag { get; set; }
        public string Note { get; set; }
    }

    public List<TodoTask> tasks = new List<TodoTask>();

    public App_TodoList(MainForm mainForm, string listType, string transferLabel) {
        this.parentForm = mainForm;
        this.listType = listType;
        this.transferLabel = transferLabel;
        this.dataFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"todo_{listType}.txt");

        this.BackColor = iosBackground;
        this.Padding = new Padding(15);
        this.AutoScaleMode = AutoScaleMode.Dpi; // 核心：支援高 DPI 縮放

        InitializeUI();
        LoadTasks();
    }

    private void InitializeUI() {
        // --- 頂部標題與新增按鈕 ---
        TableLayoutPanel header = new TableLayoutPanel() { 
            Dock = DockStyle.Top, 
            Height = 50, 
            ColumnCount = 2 
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));

        string titleText = listType == "todo" ? "待辦事項" : "待規事項";
        Label lblTitle = new Label() { 
            Text = titleText, 
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
        btnAdd.Click += (s, e) => { new EditTodoTaskWindow(this, -1, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);
        this.Controls.Add(header);

        // --- 任務卡片容器 (支援拖曳排序) ---
        taskPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            AutoScroll = true, 
            FlowDirection = FlowDirection.TopDown, 
            WrapContents = false, 
            BackColor = iosBackground, 
            AllowDrop = true,
            Padding = new Padding(0, 10, 0, 0)
        };
        
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
    }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 25 : 450;
        
        foreach (var t in tasks) {
            // 卡片背景
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
                ColumnCount = 6, 
                RowCount = 1, 
                AutoSize = true, 
                Padding = new Padding(10) 
            };
            
            // 欄位比例設定：刪除 | 轉移 | 備註 | 內容(自動延展) | 標籤 | 編輯
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f)); 

            // 1. 刪除按鈕
            Button btnDel = CreateIconButton("✕", iosRed, Color.White);
            btnDel.Click += (s, e) => { 
                if (MessageBox.Show("確定移除此任務？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK) {
                    DeleteTask(t);
                }
            };

            // 2. 轉移按鈕 (轉待辦/轉待規)
            Button btnTransfer = CreateIconButton(transferLabel, iosGreen, Color.White);
            btnTransfer.Click += (s, e) => {
                if (TargetList != null) {
                    TargetList.AddTask(t.Name, t.ColorStr, t.Tag, t.Note);
                    DeleteTask(t);
                }
            };

            // 3. 備註按鈕 (有備註顯示黃色，無則為灰色)
            bool hasNote = !string.IsNullOrWhiteSpace(t.Note);
            Button btnNote = CreateIconButton("註", hasNote ? iosYellow : Color.FromArgb(230, 230, 235), hasNote ? Color.Black : iosGrayText);
            btnNote.Click += (s, e) => {
                // 若有實作獨立的備註編輯視窗可在此呼叫，此處直接整合至編輯視窗
                new EditTodoTaskWindow(this, tasks.IndexOf(t), t).ShowDialog();
            };

            // 4. 任務文字 (支援拖曳)
            Label lbl = new Label() { 
                Text = t.Name, 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleLeft, 
                AutoSize = true, 
                Font = MainFont, 
                Padding = new Padding(10, 5, 0, 5),
                Cursor = Cursors.SizeAll 
            };
            // 依據 ColorStr 設定文字顏色
            try { lbl.ForeColor = Color.FromName(t.ColorStr); } catch { lbl.ForeColor = Color.Black; }
            if (lbl.ForeColor.Name == "0" || lbl.ForeColor == Color.Transparent) lbl.ForeColor = Color.Black;

            lbl.MouseDown += (sender, e) => {
                if (e.Button == MouseButtons.Left) card.DoDragDrop(card, DragDropEffects.Move);
            };

            // 5. 標籤顯示
            Label lblTag = new Label() {
                Text = string.IsNullOrEmpty(t.Tag) ? "" : $"[{t.Tag}]",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = iosAppleBlue,
                Font = new Font("Microsoft JhengHei UI", 9f, FontStyle.Bold)
            };

            // 6. 編輯按鈕
            Button btnEdit = CreateIconButton("修", iosAppleBlue, Color.White);
            btnEdit.Click += (s, e) => {
                new EditTodoTaskWindow(this, tasks.IndexOf(t), t).ShowDialog();
            };

            tlp.Controls.Add(btnDel, 0, 0);
            tlp.Controls.Add(btnTransfer, 1, 0);
            tlp.Controls.Add(btnNote, 2, 0);
            tlp.Controls.Add(lbl, 3, 0);
            tlp.Controls.Add(lblTag, 4, 0);
            tlp.Controls.Add(btnEdit, 5, 0);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    private Button CreateIconButton(string text, Color backColor, Color foreColor) {
        Button btn = new Button() { 
            Text = text, 
            Dock = DockStyle.Fill, 
            Height = 35, 
            BackColor = backColor, 
            ForeColor = foreColor, 
            FlatStyle = FlatStyle.Flat, 
            Cursor = Cursors.Hand, 
            Margin = new Padding(2), 
            Font = BoldFont 
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    // --- 外部呼叫 API ---
    public void AddTask(string name, string colorStr, string tag, string note) {
        tasks.Add(new TodoTask() { Name = name, ColorStr = colorStr, Tag = tag, Note = note });
        SaveTasks();
        RefreshUI();
    }

    public void DeleteTask(TodoTask t) {
        if (tasks.Contains(t)) {
            tasks.Remove(t);
            SaveTasks();
            RefreshUI();
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
            
            var item = tasks[currentIdx];
            tasks.RemoveAt(currentIdx);
            
            int finalIdx = Math.Min(targetIdx, tasks.Count);
            tasks.Insert(finalIdx, item);

            dragInsertIndex = -1; 
            taskPanel.Invalidate(); 
            SaveTasks(); 
            RefreshUI(); 
        }
    }

    // --- 存檔與載入 ---
    public void SaveTasks() {
        List<string> lines = new List<string>();
        foreach(var t in tasks) {
            string safeNote = string.IsNullOrEmpty(t.Note) ? "" : Convert.ToBase64String(Encoding.UTF8.GetBytes(t.Note));
            lines.Add($"{t.Name}|{t.ColorStr}|{t.Tag}|{safeNote}");
        }
        File.WriteAllLines(dataFile, lines);
    }

    public void LoadTasks() {
        if(!File.Exists(dataFile)) return;
        tasks.Clear();
        foreach(var l in File.ReadAllLines(dataFile)) {
            var p = l.Split('|');
            if(p.Length >= 4) {
                string note = "";
                try { note = string.IsNullOrEmpty(p[3]) ? "" : Encoding.UTF8.GetString(Convert.FromBase64String(p[3])); } catch { note = p[3]; }
                tasks.Add(new TodoTask() { Name = p[0], ColorStr = p[1], Tag = p[2], Note = note });
            }
        }
        RefreshUI();
    }
}

// ==========================================
// 視窗：新增/編輯任務 (iOS 風格)
// ==========================================
public class EditTodoTaskWindow : Form {
    private App_TodoList parent;
    private int index;
    private TextBox txtName, txtTag, txtNote;
    private ComboBox cbColor;

    public EditTodoTaskWindow(App_TodoList p, int idx, App_TodoList.TodoTask item) {
        this.parent = p;
        this.index = idx;
        this.Text = idx == -1 ? "新增任務" : "編輯任務";
        this.Width = 450; 
        this.Height = 450;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false; 
        this.MinimizeBox = false;
        this.BackColor = Color.White; 
        this.AutoScaleMode = AutoScaleMode.Dpi; 
        this.Font = new Font("Microsoft JhengHei UI", 10.5f);

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(25) };
        
        // 名稱
        f.Controls.Add(new Label() { Text = "任務名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5), ForeColor = Color.FromArgb(142, 142, 147) });
        txtName = new TextBox() { Width = 380, Text = item?.Name ?? "", Margin = new Padding(0, 0, 0, 15), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtName);

        // 標籤與顏色區塊 (並排)
        TableLayoutPanel row2 = new TableLayoutPanel() { Width = 380, Height = 60, ColumnCount = 2, Margin = new Padding(0, 0, 0, 15) };
        row2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        row2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        
        Panel pnlTag = new Panel() { Dock = DockStyle.Fill };
        pnlTag.Controls.Add(new Label() { Text = "分類標籤：", AutoSize = true, ForeColor = Color.FromArgb(142, 142, 147), Location = new Point(0, 0) });
        txtTag = new TextBox() { Width = 170, Text = item?.Tag ?? "", Location = new Point(0, 25), BorderStyle = BorderStyle.FixedSingle };
        pnlTag.Controls.Add(txtTag);

        Panel pnlColor = new Panel() { Dock = DockStyle.Fill };
        pnlColor.Controls.Add(new Label() { Text = "文字顏色：", AutoSize = true, ForeColor = Color.FromArgb(142, 142, 147), Location = new Point(0, 0) });
        cbColor = new ComboBox() { Width = 170, DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(0, 25) };
        cbColor.Items.AddRange(new string[] { "Black", "Red", "Blue", "Green", "Orange", "Purple" });
        cbColor.SelectedItem = string.IsNullOrEmpty(item?.ColorStr) ? "Black" : item.ColorStr;
        pnlColor.Controls.Add(cbColor);

        row2.Controls.Add(pnlTag, 0, 0);
        row2.Controls.Add(pnlColor, 1, 0);
        f.Controls.Add(row2);

        // 備註
        f.Controls.Add(new Label() { Text = "詳細備註：", AutoSize = true, Margin = new Padding(0, 0, 0, 5), ForeColor = Color.FromArgb(142, 142, 147) });
        txtNote = new TextBox() { Width = 380, Height = 100, Multiline = true, ScrollBars = ScrollBars.Vertical, Text = item?.Note ?? "", Margin = new Padding(0, 0, 0, 20), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtNote);

        // 儲存按鈕
        Button btnSave = new Button() { 
            Text = "儲存任務", Width = 380, Height = 42, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold)
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtName.Text)) {
                MessageBox.Show("任務名稱不可為空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (index == -1) {
                parent.AddTask(txtName.Text, cbColor.Text, txtTag.Text, txtNote.Text);
            } else {
                parent.tasks[index].Name = txtName.Text;
                parent.tasks[index].ColorStr = cbColor.Text;
                parent.tasks[index].Tag = txtTag.Text;
                parent.tasks[index].Note = txtNote.Text;
                parent.SaveTasks();
                parent.RefreshUI();
            }
            this.Close();
        };
        f.Controls.Add(btnSave);
        this.Controls.Add(f);
    }
}
