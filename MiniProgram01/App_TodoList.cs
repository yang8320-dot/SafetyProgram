using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Threading;

public class App_TodoList : UserControl {
    private string activeFile;
    private string addedLog;
    private string doneLog;

    public App_TodoList TargetList { get; set; }
    private string moveBtnText;

    private TextBox inputField;
    private FlowLayoutPanel taskContainer;

    // 將任務資料擴充，包含時間、顏色、以及「詳細說明 (Note)」
    public class TaskInfo {
        public DateTime Time;
        public string Color;
        public string Note;
        public TaskInfo(DateTime t, string c, string n) { Time = t; Color = c; Note = n; }
    }
    private Dictionary<string, TaskInfo> taskData = new Dictionary<string, TaskInfo>();
    
    private int dragInsertIndex = -1; 
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);
    private MainForm mainForm;

    private readonly string[] colorCycle = { "Black", "Red", "DodgerBlue", "MediumOrchid", "DarkGreen", "DarkOrange" };

    public App_TodoList(MainForm parent, string filePrefix, string moveText) {
        this.mainForm = parent; 
        this.moveBtnText = moveText;
        this.Dock = DockStyle.Fill; 
        
        activeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePrefix + "_active.txt");
        addedLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePrefix + "_history_added.txt");
        doneLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePrefix + "_history_completed.txt");

        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // 頂部輸入區
        TableLayoutPanel top = new TableLayoutPanel();
        top.Dock = DockStyle.Top;
        top.Height = 40;
        top.ColumnCount = 2;
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75f)); 

        inputField = new TextBox();
        inputField.Dock = DockStyle.Fill;
        inputField.Font = MainFont;
        inputField.Margin = new Padding(0, 5, 5, 0);
        inputField.KeyDown += new KeyEventHandler(InputField_KeyDown);
        
        Button btnAdd = new Button();
        btnAdd.Text = "新增";
        btnAdd.Dock = DockStyle.Fill;
        btnAdd.FlatStyle = FlatStyle.Flat;
        btnAdd.BackColor = AppleBlue;
        btnAdd.ForeColor = Color.White;
        btnAdd.Font = new Font(MainFont.FontFamily, 10f, FontStyle.Bold);
        btnAdd.Margin = new Padding(0, 3, 0, 5);
        btnAdd.Click += new EventHandler(BtnAdd_Click);

        top.Controls.Add(inputField, 0, 0);
        top.Controls.Add(btnAdd, 1, 0);

        // 任務清單容器
        taskContainer = new FlowLayoutPanel();
        taskContainer.Dock = DockStyle.Fill;
        taskContainer.AutoScroll = true;
        taskContainer.FlowDirection = FlowDirection.TopDown;
        taskContainer.WrapContents = false;
        taskContainer.BackColor = Color.White;
        taskContainer.AllowDrop = true;
        
        taskContainer.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskContainer.DragOver += OnTaskDragOver;
        taskContainer.DragLeave += (s, e) => { dragInsertIndex = -1; taskContainer.Invalidate(); };
        taskContainer.DragDrop += OnTaskDragDrop;
        taskContainer.Paint += OnTaskContainerPaint;

        taskContainer.Resize += (s, e) => {
            int safeWidth = taskContainer.ClientSize.Width - 25; 
            if (safeWidth > 0) {
                foreach (Control c in taskContainer.Controls) {
                    if (c is TableLayoutPanel) c.Width = safeWidth;
                }
            }
        };

        this.Controls.Add(top);
        this.Controls.Add(taskContainer);
        taskContainer.BringToFront(); 
        
        LoadTasks();
    }

    private void InputField_KeyDown(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) { 
            e.SuppressKeyPress = true; 
            AddTask(inputField.Text); 
            inputField.Text = ""; 
        }
    }

    private void BtnAdd_Click(object sender, EventArgs e) {
        AddTask(inputField.Text); 
        inputField.Text = "";
    }

    public void AddTask(string text, string colorName = "Black", string source = "手動", string note = "") {
        text = text.Trim(); 
        if (string.IsNullOrEmpty(text) || taskData.ContainsKey(text)) return;
        
        DateTime now = DateTime.Now;
        taskData[text] = new TaskInfo(now, colorName, note);
        
        CreateTaskUI(text, colorName);
        SafeAppendLog(addedLog, string.Format("[{0}] {1}: {2}\n", now.ToString("yyyy-MM-dd HH:mm"), source, text));
        SaveActive();
    }

    private void CreateTaskUI(string text, string textColorName) {
        Color textColor = Color.FromName(textColorName);
        int startWidth = taskContainer.ClientSize.Width > 50 ? taskContainer.ClientSize.Width - 25 : 450;

        TableLayoutPanel item = new TableLayoutPanel();
        item.Width = startWidth;
        item.AutoSize = true;
        item.Margin = new Padding(5, 3, 5, 8);
        item.BackColor = Color.White;
        item.ColumnCount = 6;
        item.RowCount = 1;
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30f)); 
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); // 註
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65f)); // 轉移
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); // 色
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 35f)); // 修
        
        CheckBox chk = new CheckBox();
        chk.Dock = DockStyle.Fill;
        chk.Cursor = Cursors.Hand;
        chk.BackColor = Color.Transparent;
        chk.ForeColor = textColor;
        chk.Padding = new Padding(5, 5, 0, 0);
        chk.CheckedChanged += (s, e) => {
            if (chk.Checked) {
                if (taskData.ContainsKey(text)) {
                    SafeAppendLog(doneLog, string.Format("[完成:{0}] {1} (建立於:{2})\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), text, taskData[text].Time.ToString("yyyy-MM-dd HH:mm")));
                    taskData.Remove(text);
                }
                taskContainer.Controls.Remove(item);
                SaveActive();
            }
        };

        Label lbl = new Label();
        lbl.Text = text;
        lbl.Dock = DockStyle.Fill;
        lbl.Font = MainFont;
        lbl.ForeColor = textColor;
        lbl.AutoSize = true;
        lbl.Padding = new Padding(0, 5, 0, 5);
        lbl.Cursor = Cursors.SizeAll;
        lbl.BackColor = Color.Transparent;
        lbl.TextAlign = ContentAlignment.MiddleLeft;

        // ==========================================
        // 【核心機制】：註解按鈕與變色邏輯
        // ==========================================
        Button btnNote = new Button();
        btnNote.Text = "註";
        btnNote.Dock = DockStyle.Fill;
        btnNote.Height = 28;
        btnNote.FlatStyle = FlatStyle.Flat;
        btnNote.Cursor = Cursors.Hand;
        btnNote.FlatAppearance.BorderSize = 0;
        btnNote.Margin = new Padding(2);
        btnNote.Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold);

        // 定義一個變色方法，隨時檢查是否有文字
        Action updateNoteStyle = () => {
            if (taskData.ContainsKey(text)) {
                if (!string.IsNullOrEmpty(taskData[text].Note)) {
                    // ★ 有文字時：變成醒目的金黃色 ★
                    btnNote.BackColor = Color.FromArgb(255, 193, 7); 
                    btnNote.ForeColor = Color.Black;
                } else {
                    // ★ 沒文字時：保持低調淺灰色 ★
                    btnNote.BackColor = Color.FromArgb(235, 235, 235); 
                    btnNote.ForeColor = Color.FromName(taskData[text].Color); 
                }
            }
        };
        
        // 建立 UI 時先檢查一次變色
        updateNoteStyle();

        btnNote.Click += (s, e) => {
            string currentNote = taskData.ContainsKey(text) ? taskData[text].Note : "";
            string newNote = ShowNoteEditBox(text, currentNote);
            if (newNote != null && taskData.ContainsKey(text)) {
                taskData[text].Note = newNote;
                updateNoteStyle(); // 編輯完馬上刷新按鈕顏色！
                SaveActive();
            }
        };

        Button btnMove = new Button();
        btnMove.Text = moveBtnText;
        btnMove.Dock = DockStyle.Fill;
        btnMove.Height = 28;
        btnMove.FlatStyle = FlatStyle.Flat;
        btnMove.BackColor = Color.FromArgb(235, 230, 255);
        btnMove.ForeColor = textColor;
        btnMove.Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold);
        btnMove.Cursor = Cursors.Hand;
        btnMove.FlatAppearance.BorderSize = 0;
        btnMove.Margin = new Padding(2);

        Button btnColor = new Button();
        btnColor.Text = "色";
        btnColor.Dock = DockStyle.Fill;
        btnColor.Height = 28;
        btnColor.FlatStyle = FlatStyle.Flat;
        btnColor.BackColor = Color.FromArgb(211, 227, 253);
        btnColor.ForeColor = textColor;
        btnColor.Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold);
        btnColor.Cursor = Cursors.Hand;
        btnColor.FlatAppearance.BorderSize = 0;
        btnColor.Margin = new Padding(2);

        Button btnEdit = new Button();
        btnEdit.Text = "修";
        btnEdit.Dock = DockStyle.Fill;
        btnEdit.Height = 28;
        btnEdit.FlatStyle = FlatStyle.Flat;
        btnEdit.BackColor = Color.FromArgb(255, 210, 210);
        btnEdit.ForeColor = textColor;
        btnEdit.Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold);
        btnEdit.Cursor = Cursors.Hand;
        btnEdit.FlatAppearance.BorderSize = 0;
        btnEdit.Margin = new Padding(2);

        btnMove.Click += (s, e) => {
            if (TargetList != null) {
                string currentColorName = taskData[text].Color;
                string currentNote = taskData[text].Note;
                TargetList.AddTask(text, currentColorName, "轉移寫入", currentNote); 
                if (taskData.ContainsKey(text)) {
                    SafeAppendLog(doneLog, string.Format("[轉出至{0}:{1}] {2}\n", moveBtnText.Replace("轉",""), DateTime.Now.ToString("yyyy-MM-dd HH:mm"), text));
                    taskData.Remove(text);
                }
                taskContainer.Controls.Remove(item);
                SaveActive();
            }
        };

        btnColor.Click += (s, e) => {
            string currentColorName = taskData[text].Color;
            int nextIdx = (Array.IndexOf(colorCycle, currentColorName) + 1) % colorCycle.Length;
            string nextColorName = colorCycle[nextIdx];
            
            taskData[text].Color = nextColorName;
            Color newTextColor = Color.FromName(nextColorName);
            
            lbl.ForeColor = newTextColor;
            chk.ForeColor = newTextColor;
            btnEdit.ForeColor = newTextColor;
            btnColor.ForeColor = newTextColor;
            btnMove.ForeColor = newTextColor; 
            updateNoteStyle();
            
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

        item.Controls.Add(chk, 0, 0);
        item.Controls.Add(lbl, 1, 0);
        item.Controls.Add(btnNote, 2, 0);
        item.Controls.Add(btnMove, 3, 0);
        item.Controls.Add(btnColor, 4, 0);
        item.Controls.Add(btnEdit, 5, 0);
        
        taskContainer.Controls.Add(item); 
        taskContainer.Controls.SetChildIndex(item, 0);
    }

    private void SaveActive() {
        ThreadPool.QueueUserWorkItem(_ => {
            List<string> lines = new List<string>();
            try {
                this.Invoke(new Action(() => {
                    foreach (Control ctrl in taskContainer.Controls) {
                        if (ctrl is TableLayoutPanel p) {
                            foreach (Control sub in p.Controls) {
                                if (sub is Label lbl && taskData.ContainsKey(lbl.Text)) { 
                                    TaskInfo info = taskData[lbl.Text];
                                    string encodedNote = EncodeBase64(info.Note);
                                    lines.Add(string.Format("{0}|{1}|{2}|{3}", lbl.Text, info.Time.ToString(), info.Color, encodedNote)); 
                                    break; 
                                }
                            }
                        }
                    }
                }));

                for (int i = 0; i < 5; i++) {
                    try { File.WriteAllLines(activeFile, lines); return; } 
                    catch (IOException) { Thread.Sleep(150); }
                }
            } catch { }
        });
    }

    private void SafeAppendLog(string path, string content) {
        for (int i = 0; i < 5; i++) {
            try { File.AppendAllText(path, content); break; }
            catch (IOException) { Thread.Sleep(150); }
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
                    string note = p.Length >= 4 ? DecodeBase64(p[3]) : "";
                    
                    taskData[text] = new TaskInfo(time, color, note);
                    CreateTaskUI(text, color); 
                }
            }
        } catch { }
    }

    private void OnTaskDragOver(object sender, DragEventArgs e) {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskContainer.PointToClient(new Point(e.X, e.Y));
        Control target = taskContainer.GetChildAtPoint(clientPoint);
        if (target != null) {
            if (!(target is TableLayoutPanel)) target = target.Parent;
            int idx = taskContainer.Controls.GetChildIndex(target);
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            if (dragInsertIndex != idx) { dragInsertIndex = idx; taskContainer.Invalidate(); }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e) {
        if (dragInsertIndex != -1 && taskContainer.Controls.Count > 0) {
            int y = (dragInsertIndex < taskContainer.Controls.Count) ? taskContainer.Controls[dragInsertIndex].Top - 2 : taskContainer.Controls[taskContainer.Controls.Count - 1].Bottom + 2;
            e.Graphics.FillRectangle(new SolidBrush(AppleBlue), 5, y, taskContainer.Width - 30, 3);
        }
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        TableLayoutPanel draggedItem = (TableLayoutPanel)e.Data.GetData(typeof(TableLayoutPanel));
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
        
        form.Controls.AddRange(new Control[] { lbl, txt, btnOk });
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) return txt.Text.Trim();
        return "";
    }

    // 顯示詳細說明 (註) 的輸入視窗
    private string ShowNoteEditBox(string taskName, string currentNote) {
        Form form = new Form() { Width = 400, Height = 350, Text = "任務詳細說明 (註)", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
        
        Label lbl = new Label() { Text = "【 " + taskName + " 】", Left = 15, Top = 15, Width = 350, Height = 45, Font = new Font(MainFont, FontStyle.Bold), ForeColor = AppleBlue };
        TextBox txt = new TextBox() { Left = 15, Top = 65, Width = 350, Height = 180, Multiline = true, WordWrap = true, ScrollBars = ScrollBars.Vertical, Font = MainFont, Text = currentNote };
        Button btnOk = new Button() { Text = "儲存說明", Left = 265, Top = 260, Width = 100, Height = 35, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        
        form.Controls.AddRange(new Control[] { lbl, txt, btnOk });
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) return txt.Text.Trim();
        return null;
    }

    // 將多行文字進行安全的 Base64 編碼與解碼
    private string EncodeBase64(string text) {
        if (string.IsNullOrEmpty(text)) return "";
        return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(text));
    }

    private string DecodeBase64(string base64) {
        if (string.IsNullOrEmpty(base64)) return "";
        try { return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(base64)); }
        catch { return base64; } 
    }
}
