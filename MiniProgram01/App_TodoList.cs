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
        
        taskContainer.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskContainer.DragOver += (s, e) => {
            Point p = taskContainer.PointToClient(new Point(e.X, e.Y));
            Control target = taskContainer.GetChildAtPoint(p);
            if (target != null) e.Effect = DragDropEffects.Move;
        };
        taskContainer.DragDrop += OnTaskDragDrop;

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
            Width = 330, AutoSize = true, Padding = new Padding(5), Margin = new Padding(0, 0, 0, 2), BackColor = Color.FromArgb(252, 252, 254)
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

        Label lbl = new Label() { 
            Text = text, Dock = DockStyle.Fill, Font = MainFont, AutoSize = true, MaximumSize = new Size(280, 0), Padding = new Padding(0, 5, 0, 5), Cursor = Cursors.SizeAll 
        };

        lbl.MouseDoubleClick += (s, e) => {
            string oldText = lbl.Text;
            string newText = App_Shortcuts.ShowInputBox("修改任務內容：", "✏️ 編輯任務", oldText);
            if (!string.IsNullOrEmpty(newText) && newText != oldText && !taskDates.ContainsKey(newText)) {
                taskDates[newText] = taskDates[oldText]; taskDates.Remove(oldText);
                lbl.Text = newText; text = newText; SaveActive();
            }
        };

        lbl.MouseDown += (s, e) => {
            if (e.Button == MouseButtons.Left) item.DoDragDrop(item, DragDropEffects.Move);
        };

        item.Controls.Add(lbl); item.Controls.Add(chk);
        taskContainer.Controls.Add(item);
        taskContainer.Controls.SetChildIndex(item, 0);
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        Panel draggedItem = (Panel)e.Data.GetData(typeof(Panel));
        Point clientPoint = taskContainer.PointToClient(new Point(e.X, e.Y));
        Control target = taskContainer.GetChildAtPoint(clientPoint);
        if (target != null && draggedItem != null) {
            if (!(target is Panel)) target = target.Parent;
            int targetIdx = taskContainer.Controls.GetChildIndex(target);
            taskContainer.Controls.SetChildIndex(draggedItem, targetIdx);
            SaveActive();
        }
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
