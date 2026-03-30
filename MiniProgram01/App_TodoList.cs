using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

public class App_TodoList : UserControl {
    private string activeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_active.txt");
    private string addedLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_added.txt");
    private string doneLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_completed.txt");

    private TextBox inputField;
    private FlowLayoutPanel taskContainer;
    private Dictionary<string, DateTime> taskDates = new Dictionary<string, DateTime>();
    
    // 拖曳視覺輔助變數
    private int dragInsertIndex = -1; 
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);

    public App_TodoList(MainForm parent) {
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        Panel top = new Panel() { Dock = DockStyle.Top, Height = 40 };
        inputField = new TextBox() { Width = 240, Font = MainFont, Location = new Point(0, 5) };
        inputField.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AddTask(inputField.Text); } };
        
        Button btnAdd = new Button() { 
            Text = "新增", Left = 250, Top = 3, Width = 65, Height = 30, 
            FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont, FontStyle.Bold) 
        };
        btnAdd.Click += (s, e) => AddTask(inputField.Text);
        top.Controls.AddRange(new Control[] { inputField, btnAdd });

        taskContainer = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            AutoScroll = true, 
            FlowDirection = FlowDirection.TopDown, 
            WrapContents = false, 
            BackColor = Color.White,
            AllowDrop = true 
        };
        
        // --- 拖曳邏輯與分隔線繪製 ---
        taskContainer.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskContainer.DragOver += OnTaskDragOver;
        taskContainer.DragLeave += (s, e) => { dragInsertIndex = -1; taskContainer.Invalidate(); };
        taskContainer.DragDrop += OnTaskDragDrop;
        taskContainer.Paint += OnTaskContainerPaint; // 繪製藍色分隔線

        this.Controls.Add(taskContainer);
        this.Controls.Add(top);
        
        LoadTasks();
    }

    public void AddTaskExternally(string text) { 
        if (!taskDates.ContainsKey(text)) AddTask(text, true); 
    }

    private void AddTask(string text, bool auto = false) {
        text = text.Trim(); 
        if (string.IsNullOrEmpty(text) || taskDates.ContainsKey(text)) return;
        DateTime now = DateTime.Now;
        taskDates[text] = now;
        CreateTaskUI(text);
        File.AppendAllText(addedLog, string.Format("[{0}] {1}: {2}\n", now.ToString("yyyy-MM-dd HH:mm"), auto ? "排程" : "手動", text));
        SaveActive();
        inputField.Text = "";
    }

    private void CreateTaskUI(string text) {
        Panel item = new Panel() { 
            Width = 335, AutoSize = true, Padding = new Padding(5), Margin = new Padding(0, 0, 0, 2), BackColor = Color.FromArgb(252, 252, 254)
        };

        CheckBox chk = new CheckBox() { Dock = DockStyle.Left, Width = 30, Cursor = Cursors.Hand };
        chk.CheckedChanged += (s, e) => {
            if (chk.Checked) {
                if (taskDates.ContainsKey(text)) {
                    File.AppendAllText(doneLog, string.Format("[完成:{0}] {1} (建立於:{2})\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), text, taskDates[text].ToString("yyyy-MM-dd HH:mm")));
                    taskDates.Remove(text);
                }
                taskContainer.Controls.Remove(item);
                SaveActive();
            }
        };

        // 「修」按鈕：放在最後面
        Button btnEdit = new Button() { 
            Text = "修", Dock = DockStyle.Right, Width = 35, FlatStyle = FlatStyle.Flat, 
            ForeColor = AppleBlue, Font = new Font(MainFont.FontFamily, 9f), Cursor = Cursors.Hand 
        };
        btnEdit.FlatAppearance.BorderSize = 0;

        Label lbl = new Label() { 
            Text = text, Dock = DockStyle.Fill, Font = MainFont, AutoSize = true, 
            MaximumSize = new Size(260, 0), Padding = new Padding(0, 5, 0, 5), Cursor = Cursors.SizeAll 
        };

        // 統一編輯邏輯
        Action triggerEdit = () => {
            string oldText = lbl.Text;
            string newText = ShowLargeEditBox(oldText); // 呼叫大尺寸編輯框
            if (!string.IsNullOrEmpty(newText) && newText != oldText && !taskDates.ContainsKey(newText)) {
                taskDates[newText] = taskDates[oldText]; taskDates.Remove(oldText);
                lbl.Text = newText; text = newText; SaveActive();
            }
        };

        lbl.MouseDoubleClick += (s, e) => triggerEdit();
        btnEdit.Click += (s, e) => triggerEdit();

        lbl.MouseDown += (s, e) => {
            if (e.Button == MouseButtons.Left) item.DoDragDrop(item, DragDropEffects.Move);
        };

        item.Controls.Add(lbl);
        item.Controls.Add(chk);
        item.Controls.Add(btnEdit);
        
        taskContainer.Controls.Add(item);
        taskContainer.Controls.SetChildIndex(item, 0);
    }

    // --- 拖曳分隔線繪製邏輯 ---
    private void OnTaskDragOver(object sender, DragEventArgs e) {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskContainer.PointToClient(new Point(e.X, e.Y));
        Control target = taskContainer.GetChildAtPoint(clientPoint);
        
        if (target != null) {
            if (!(target is Panel)) target = target.Parent;
            int idx = taskContainer.Controls.GetChildIndex(target);
            
            // 決定畫在目標的上方還是下方
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            
            if (dragInsertIndex != idx) {
                dragInsertIndex = idx;
                taskContainer.Invalidate(); // 重新繪製線條
            }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e) {
        if (dragInsertIndex != -1 && taskContainer.Controls.Count > 0) {
            int y = 0;
            if (dragInsertIndex < taskContainer.Controls.Count) {
                y = taskContainer.Controls[dragInsertIndex].Top - 2;
            } else {
                y = taskContainer.Controls[taskContainer.Controls.Count - 1].Bottom + 2;
            }
            // 畫出顯眼的藍色分隔線
            e.Graphics.FillRectangle(new SolidBrush(AppleBlue), 5, y, taskContainer.Width - 25, 3);
        }
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        Panel draggedItem = (Panel)e.Data.GetData(typeof(Panel));
        if (draggedItem != null && dragInsertIndex != -1) {
            int targetIdx = dragInsertIndex;
            int currentIdx = taskContainer.Controls.GetChildIndex(draggedItem);
            
            if (currentIdx < targetIdx) targetIdx--; // 修正因移除自身造成的索引偏移
            
            taskContainer.Controls.SetChildIndex(draggedItem, targetIdx);
            dragInsertIndex = -1;
            taskContainer.Invalidate();
            SaveActive();
        }
    }

    // --- 大尺寸編輯視窗 ---
    private string ShowLargeEditBox(string defaultValue) {
        Form form = new Form() { 
            Width = 450, Height = 250, Text = "✏️ 滾動式編輯任務", 
            StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false, MinimizeBox = false 
        };
        
        Label lbl = new Label() { Text = "請修正任務內容：", Left = 15, Top = 15, AutoSize = true, Font = MainFont };
        
        TextBox txt = new TextBox() { 
            Left = 15, Top = 45, Width = 405, Height = 100, 
            Multiline = true, WordWrap = true, ScrollBars = ScrollBars.Vertical,
            Font = MainFont, Text = defaultValue 
        };
        
        Button btnOk = new Button() { 
            Text = "確認修改", Left = 320, Top = 165, Width = 100, Height = 35, 
            DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White 
        };
        
        form.Controls.AddRange(new Control[] { lbl, txt, btnOk });
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; // 游標移到最後

        return (form.ShowDialog() == DialogResult.OK) ? txt.Text.Trim() : "";
    }

    private void SaveActive() {
        List<string> lines = new List<string>();
        foreach (Control ctrl in taskContainer.Controls) {
            if (ctrl is Panel p) {
                foreach (Control sub in p.Controls) {
                    if (sub is Label lbl) { lines.Add(lbl.Text + "|" + taskDates[lbl.Text].ToString()); break; }
                }
            }
        }
        File.WriteAllLines(activeFile, lines);
    }

    private void LoadTasks() {
        if (!File.Exists(activeFile)) return;
        string[] lines = File.ReadAllLines(activeFile);
        Array.Reverse(lines); 
        foreach (string l in lines) {
            string[] p = l.Split('|');
            if (p.Length >= 2) { taskDates[p[0]] = DateTime.Parse(p[1]); CreateTaskUI(p[0]); }
        }
    }
}
