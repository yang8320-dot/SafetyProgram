using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Threading;

public class App_TodoList : UserControl {
    private string activeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_active.txt");
    private string addedLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_added.txt");
    private string doneLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_completed.txt");

    private TextBox inputField;
    private FlowLayoutPanel taskContainer;
    private Dictionary<string, Tuple<DateTime, string>> taskData = new Dictionary<string, Tuple<DateTime, string>>();
    
    private int dragInsertIndex = -1; 
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);

    private MainForm mainForm;

    // 【指定順序】文字顏色循環：黑(預設) -> 亮紅 -> 亮藍 -> 亮紫 -> 深綠 -> 深橘
    private readonly string[] colorCycle = { "Black", "Red", "DodgerBlue", "MediumOrchid", "DarkGreen", "DarkOrange" };

    public App_TodoList(MainForm parent) {
        this.mainForm = parent; 
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

        taskContainer = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White, AllowDrop = true };
        
        taskContainer.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskContainer.DragOver += OnTaskDragOver;
        taskContainer.DragLeave += (s, e) => { dragInsertIndex = -1; taskContainer.Invalidate(); };
        taskContainer.DragDrop += OnTaskDragDrop;
        taskContainer.Paint += OnTaskContainerPaint;

        this.Controls.Add(taskContainer);
        this.Controls.Add(top);
        
        LoadTasks();
    }

    public void AddTask(string text, bool auto = false) {
        text = text.Trim(); 
        if (string.IsNullOrEmpty(text) || taskData.ContainsKey(text)) return;
        
        DateTime now = DateTime.Now;
        string defaultTextColorName = "Black"; 
        taskData[text] = new Tuple<DateTime, string>(now, defaultTextColorName);
        
        CreateTaskUI(text, defaultTextColorName);
        SafeAppendLog(addedLog, string.Format("[{0}] {1}: {2}\n", now.ToString("yyyy-MM-dd HH:mm"), auto ? "排程" : "手動", text));
        SaveActive();
        inputField.Text = "";
    }

    private void CreateTaskUI(string text, string textColorName) {
        Color textColor = Color.FromName(textColorName);
        // 【修正】框 (背景) 固定為白色
        Panel item = new Panel() { Width = 335, AutoSize = true, Padding = new Padding(5), Margin = new Padding(0, 0, 0, 2), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
        
        CheckBox chk = new CheckBox() { Dock = DockStyle.Left, Width = 30, Cursor = Cursors.Hand, BackColor = Color.Transparent, ForeColor = textColor };
        chk.CheckedChanged += (s, e) => {
            if (chk.Checked) {
                if (taskData.ContainsKey(text)) {
                    SafeAppendLog(doneLog, string.Format("[完成:{0}] {1} (建立於:{2})\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), text, taskData[text].Item1.ToString("yyyy-MM-dd HH:mm")));
                    taskData.Remove(text);
                }
                taskContainer.Controls.Remove(item);
                SaveActive();
            }
        };

        Button btnEdit = new Button() { Text = "修", Dock = DockStyle.Right, Width = 30, FlatStyle = FlatStyle.Flat, ForeColor = textColor, Font = new Font(MainFont.FontFamily, 9f), Cursor = Cursors.Hand };
        btnEdit.FlatAppearance.BorderSize = 0;

        Button btnColor = new Button() { Text = "色", Dock = DockStyle.Right, Width = 30, FlatStyle = FlatStyle.Flat, ForeColor = textColor, Font = new Font(MainFont.FontFamily, 9f), Cursor = Cursors.Hand };
        btnColor.FlatAppearance.BorderSize = 0;

        Label lbl = new Label() { Text = text, Dock = DockStyle.Fill, Font = MainFont, ForeColor = textColor, AutoSize = true, MaximumSize = new Size(230, 0), Padding = new Padding(0, 5, 0, 5), Cursor = Cursors.SizeAll, BackColor = Color.Transparent };

        // 點擊「色」按鈕：變更文字顏色，框不變
        btnColor.Click += (s, e) => {
            string currentColorName = taskData[text].Item2;
            int nextIdx = (Array.IndexOf(colorCycle, currentColorName) + 1) % colorCycle.Length;
            string nextColorName = colorCycle[nextIdx];
            
            taskData[text] = new Tuple<DateTime, string>(taskData[text].Item1, nextColorName);
            Color newTextColor = Color.FromName(nextColorName);
            
            lbl.ForeColor = newTextColor;
            chk.ForeColor = newTextColor;
            btnEdit.ForeColor = newTextColor;
            btnColor.ForeColor = newTextColor;
            
            SaveActive();
        };

        Action triggerEdit = () => {
            string oldText = lbl.Text;
            string newText = ShowLargeEditBox(oldText); 
            if (!string.IsNullOrEmpty(newText) && newText != oldText && !taskData.ContainsKey(newText)) {
                taskData[newText] = taskData[oldText]; taskData.Remove(oldText);
                lbl.Text = newText; text = newText; SaveActive();
            }
        };

        lbl.MouseDoubleClick += (s, e) => triggerEdit();
        btnEdit.Click += (s, e) => triggerEdit();
        lbl.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) item.DoDragDrop(item, DragDropEffects.Move); };

        item.Controls.Add(lbl); item.Controls.Add(chk); item.Controls.Add(btnColor); item.Controls.Add(btnEdit);
        taskContainer.Controls.Add(item); taskContainer.Controls.SetChildIndex(item, 0);
    }

    // ==========================================
    // 🛡️ 檔案存取修復邏輯 (防止「檔案正在使用」錯誤)
    // ==========================================
    private void SaveActive() {
        // 使用背景執行緒或重試機制來防止 IO 衝突
        ThreadPool.QueueUserWorkItem(_ => {
            List<string> lines = new List<string>();
            this.Invoke(new Action(() => {
                foreach (Control ctrl in taskContainer.Controls) {
                    if (ctrl is Panel p) {
                        foreach (Control sub in p.Controls) {
                            if (sub is Label lbl) { 
                                string colorName = taskData.ContainsKey(lbl.Text) ? taskData[lbl.Text].Item2 : "Black";
                                lines.Add(string.Format("{0}|{1}|{2}", lbl.Text, taskData[lbl.Text].Item1.ToString(), colorName)); 
                                break; 
                            }
                        }
                    }
                }
            }));

            // 重試機制：如果檔案被鎖定，嘗試 3 次
            for (int i = 0; i < 3; i++) {
                try {
                    File.WriteAllLines(activeFile, lines);
                    break; 
                } catch (IOException) { Thread.Sleep(200); }
            }
        });
    }

    private void SafeAppendLog(string path, string content) {
        for (int i = 0; i < 3; i++) {
            try { File.AppendAllText(path, content); break; }
            catch (IOException) { Thread.Sleep(200); }
        }
    }

    private void LoadTasks() {
        if (!File.Exists(activeFile)) return;
        try {
            string[] lines = File.ReadAllLines(activeFile);
            Array.Reverse(lines); 
            foreach (string l in lines) {
                string[] p = l.Split('|');
                if (p.Length >= 2) { 
                    string text = p[0];
                    DateTime time = DateTime.Parse(p[1]);
                    string color = p.Length >= 3 ? p[2] : "Black";
                    taskData[text] = new Tuple<DateTime, string>(time, color);
                    CreateTaskUI(text, color); 
                }
            }
        } catch { }
    }

    // --- 以下拖曳與編輯邏輯維持不變 ---
    private void OnTaskDragOver(object sender, DragEventArgs e) {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskContainer.PointToClient(new Point(e.X, e.Y));
        Control target = taskContainer.GetChildAtPoint(clientPoint);
        if (target != null) {
            if (!(target is Panel)) target = target.Parent;
            int idx = taskContainer.Controls.GetChildIndex(target);
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            if (dragInsertIndex != idx) { dragInsertIndex = idx; taskContainer.Invalidate(); }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e) {
        if (dragInsertIndex != -1 && taskContainer.Controls.Count > 0) {
            int y = (dragInsertIndex < taskContainer.Controls.Count) ? taskContainer.Controls[dragInsertIndex].Top - 2 : taskContainer.Controls[taskContainer.Controls.Count - 1].Bottom + 2;
            e.Graphics.FillRectangle(new SolidBrush(AppleBlue), 5, y, taskContainer.Width - 25, 3);
        }
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        Panel draggedItem = (Panel)e.Data.GetData(typeof(Panel));
        if (draggedItem != null && dragInsertIndex != -1) {
            int targetIdx = dragInsertIndex;
            int currentIdx = taskContainer.Controls.GetChildIndex(draggedItem);
            if (currentIdx < targetIdx) targetIdx--; 
            taskContainer.Controls.SetChildIndex(draggedItem, targetIdx);
            dragInsertIndex = -1; taskContainer.Invalidate(); SaveActive();
        }
    }

    private string ShowLargeEditBox(string defaultValue) {
        Form form = new Form() { Width = 450, Height = 250, Text = "滾動式編輯任務", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
        Label lbl = new Label() { Text = "請修正任務內容：", Left = 15, Top = 15, AutoSize = true, Font = MainFont };
        TextBox txt = new TextBox() { Left = 15, Top = 45, Width = 405, Height = 100, Multiline = true, WordWrap = true, ScrollBars = ScrollBars.Vertical, Font = MainFont, Text = defaultValue };
        Button btnOk = new Button() { Text = "確認修改", Left = 320, Top = 165, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        form.Controls.AddRange(new Control[] { lbl, txt, btnOk }); form.AcceptButton = btnOk; txt.SelectionStart = txt.Text.Length; 
        return (form.ShowDialog() == DialogResult.OK) ? txt.Text.Trim() : "";
    }
}
