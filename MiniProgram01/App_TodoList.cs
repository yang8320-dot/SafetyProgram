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
    private Dictionary<string, Tuple<DateTime, string>> taskData = new Dictionary<string, Tuple<DateTime, string>>();
    
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

        // 【自適應修正 1】頂部輸入區改用 TableLayoutPanel，按鈕跟輸入框會自動隨視窗放大縮小
        TableLayoutPanel top = new TableLayoutPanel();
        top.Dock = DockStyle.Top;
        top.Height = 40;
        top.ColumnCount = 2;
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // 輸入框佔滿剩餘空間
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75f)); // 按鈕固定 75 寬
        top.Padding = new Padding(0, 0, 0, 5);

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
        btnAdd.Margin = new Padding(0, 3, 0, 0);
        btnAdd.Click += new EventHandler(BtnAdd_Click);

        top.Controls.Add(inputField, 0, 0);
        top.Controls.Add(btnAdd, 1, 0);

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

        // 【自適應修正 2】當視窗或捲軸改變大小時，自動調整所有卡片的寬度！
        taskContainer.Resize += (s, e) => {
            int safeWidth = taskContainer.ClientSize.Width - 8;
            if (safeWidth > 0) {
                foreach (Control c in taskContainer.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
            }
        };

        this.Controls.Add(taskContainer);
        this.Controls.Add(top);
        top.BringToFront(); 
        
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

    public void AddTask(string text, string colorName = "Black", string source = "手動") {
        text = text.Trim(); 
        if (string.IsNullOrEmpty(text) || taskData.ContainsKey(text)) return;
        
        DateTime now = DateTime.Now;
        taskData[text] = new Tuple<DateTime, string>(now, colorName);
        
        CreateTaskUI(text, colorName);
        SafeAppendLog(addedLog, string.Format("[{0}] {1}: {2}\n", now.ToString("yyyy-MM-dd HH:mm"), source, text));
        SaveActive();
    }

    private void CreateTaskUI(string text, string textColorName) {
        Color textColor = Color.FromName(textColorName);
        
        // 【自適應修正 3】新增卡片時，直接讀取容器的安全寬度
        int startWidth = taskContainer.ClientSize.Width > 50 ? taskContainer.ClientSize.Width - 8 : 400;

        Panel item = new Panel();
        item.Width = startWidth;
        item.AutoSize = true;
        item.Padding = new Padding(5, 5, 2, 5);
        item.Margin = new Padding(0, 3, 0, 8);
        item.BackColor = Color.White;
        item.BorderStyle = BorderStyle.None;
        
        CheckBox chk = new CheckBox();
        chk.Dock = DockStyle.Left;
        chk.Width = 30;
        chk.Cursor = Cursors.Hand;
        chk.BackColor = Color.Transparent;
        chk.ForeColor = textColor;
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

        Button btnMove = new Button();
        btnMove.Text = moveBtnText;
        btnMove.Dock = DockStyle.Right;
        btnMove.Width = 55;
        btnMove.Height = 28;
        btnMove.FlatStyle = FlatStyle.Flat;
        btnMove.BackColor = Color.FromArgb(235, 230, 255);
        btnMove.ForeColor = textColor;
        btnMove.Font = new Font(MainFont.FontFamily, 8.5f, FontStyle.Bold);
        btnMove.Cursor = Cursors.Hand;
        btnMove.FlatAppearance.BorderSize = 0;
        btnMove.Margin = new Padding(2, 0, 2, 0);

        Button btnEdit = new Button();
        btnEdit.Text = "修";
        btnEdit.Dock = DockStyle.Right;
        btnEdit.Width = 32;
        btnEdit.Height = 28;
        btnEdit.FlatStyle = FlatStyle.Flat;
        btnEdit.BackColor = Color.FromArgb(255, 210, 210);
        btnEdit.ForeColor = textColor;
        btnEdit.Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold);
        btnEdit.Cursor = Cursors.Hand;
        btnEdit.FlatAppearance.BorderSize = 0;
        btnEdit.Margin = new Padding(2, 0, 2, 0);

        Button btnColor = new Button();
        btnColor.Text = "色";
        btnColor.Dock = DockStyle.Right;
        btnColor.Width = 32;
        btnColor.Height = 28;
        btnColor.FlatStyle = FlatStyle.Flat;
        btnColor.BackColor = Color.FromArgb(211, 227, 253);
        btnColor.ForeColor = textColor;
        btnColor.Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold);
        btnColor.Cursor = Cursors.Hand;
        btnColor.FlatAppearance.BorderSize = 0;
        btnColor.Margin = new Padding(2, 0, 2, 0);

        Label lbl = new Label();
        lbl.Text = text;
        lbl.Dock = DockStyle.Fill;
        lbl.Font = MainFont;
        lbl.ForeColor = textColor;
        lbl.AutoSize = true;
        lbl.MaximumSize = new Size(250, 0); // 字數太多時自動換行
        lbl.Padding = new Padding(0, 5, 0, 5);
        lbl.Cursor = Cursors.SizeAll;
        lbl.BackColor = Color.Transparent;

        btnMove.Click += (s, e) => {
            if (TargetList != null) {
                string currentColorName = taskData[text].Item2;
                TargetList.AddTask(text, currentColorName, "轉移寫入"); 
                if (taskData.ContainsKey(text)) {
                    SafeAppendLog(doneLog, string.Format("[轉出至{0}:{1}] {2}\n", moveBtnText.Replace("轉",""), DateTime.Now.ToString("yyyy-MM-dd HH:mm"), text));
                    taskData.Remove(text);
                }
                taskContainer.Controls.Remove(item);
                SaveActive();
            }
        };

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
            btnMove.ForeColor = newTextColor; 
            
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

        item.Controls.Add(btnEdit);
        item.Controls.Add(btnColor);
        item.Controls.Add(btnMove);
        item.Controls.Add(chk);
        item.Controls.Add(lbl);
        
        taskContainer.Controls.Add(item); 
        taskContainer.Controls.SetChildIndex(item, 0);
    }

    private void SaveActive() {
        ThreadPool.QueueUserWorkItem(_ => {
            List<string> lines = new List<string>();
            try {
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
                    taskData[text] = new Tuple<DateTime, string>(time, color);
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
        Form form = new Form();
        form.Width = 450;
        form.Height = 250;
        form.Text = "滾動式編輯任務";
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.MaximizeBox = false;
        form.MinimizeBox = false;
        
        Label lbl = new Label();
        lbl.Text = "請修正任務內容：";
        lbl.Left = 15;
        lbl.Top = 15;
        lbl.AutoSize = true;
        lbl.Font = MainFont;
        
        TextBox txt = new TextBox();
        txt.Left = 15;
        txt.Top = 45;
        txt.Width = 405;
        txt.Height = 100;
        txt.Multiline = true;
        txt.WordWrap = true;
        txt.ScrollBars = ScrollBars.Vertical;
        txt.Font = MainFont;
        txt.Text = defaultValue;
        
        Button btnOk = new Button();
        btnOk.Text = "確認修改";
        btnOk.Left = 320;
        btnOk.Top = 165;
        btnOk.Width = 100;
        btnOk.Height = 35;
        btnOk.DialogResult = DialogResult.OK;
        btnOk.FlatStyle = FlatStyle.Flat;
        btnOk.BackColor = AppleBlue;
        btnOk.ForeColor = Color.White;
        
        form.Controls.Add(lbl);
        form.Controls.Add(txt);
        form.Controls.Add(btnOk);
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) {
            return txt.Text.Trim();
        }
        return "";
    }
}
