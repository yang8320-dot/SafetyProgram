using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_TodoList : UserControl {
    private MainForm parentForm;
    private TextBox inputField;
    private CheckedListBox taskList;
    private Button btnAdd;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);

    private string activeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_active.txt");
    private string addedLogFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_added.txt");
    private string completedLogFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_completed.txt");

    private Dictionary<string, DateTime> taskDates = new Dictionary<string, DateTime>();

    public App_TodoList(MainForm mainForm) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 40 };
        inputField = new TextBox() { Location = new Point(5, 5), Width = 240, Font = MainFont, BorderStyle = BorderStyle.FixedSingle };
        inputField.KeyDown += new KeyEventHandler(OnInputKeyDown);

        btnAdd = new Button() { Text = "新增", Location = new Point(255, 4), Width = 65, Height = 26, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += new EventHandler(delegate { AddTask(); });

        topPanel.Controls.Add(inputField); topPanel.Controls.Add(btnAdd);
        this.Controls.Add(topPanel);

        taskList = new CheckedListBox() { Dock = DockStyle.Fill, Font = new Font(MainFont.FontFamily, 10.5f), CheckOnClick = true, BorderStyle = BorderStyle.None, BackColor = Color.White };
        taskList.ItemCheck += new ItemCheckEventHandler(OnTaskChecked);
        taskList.MouseDoubleClick += new MouseEventHandler(OnTaskDoubleClicked);
        
        this.Controls.Add(taskList); taskList.BringToFront(); 
        LoadActiveTasks();
    }

    private void LoadActiveTasks() {
        if (!File.Exists(activeFile)) return;
        foreach (string line in File.ReadAllLines(activeFile)) {
            if (string.IsNullOrWhiteSpace(line)) continue;
            string[] parts = line.Split(new char[] { '|' }, 2);
            if (parts.Length == 2) {
                DateTime addedTime;
                if (DateTime.TryParse(parts[1], out addedTime)) { taskList.Items.Add(parts[0]); taskDates[parts[0]] = addedTime; }
            }
        }
    }

    private void SaveActiveTasks() {
        List<string> lines = new List<string>();
        foreach (object item in taskList.Items) {
            string taskName = item.ToString();
            if (taskDates.ContainsKey(taskName)) lines.Add(taskName + "|" + taskDates[taskName].ToString("yyyy-MM-dd HH:mm:ss"));
        }
        File.WriteAllLines(activeFile, lines.ToArray());
    }

    private void AddTask() {
        string text = inputField.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;
        if (taskDates.ContainsKey(text)) { MessageBox.Show("已在清單中囉！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

        DateTime now = DateTime.Now; taskList.Items.Insert(0, text); taskDates[text] = now;
        File.AppendAllText(addedLogFile, string.Format("[{0}] 手動新增：{1}\r\n", now.ToString("yyyy-MM-dd HH:mm:ss"), text));
        SaveActiveTasks(); inputField.Text = ""; inputField.Focus();
    }

    private void OnInputKeyDown(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AddTask(); } }

    private void OnTaskChecked(object sender, ItemCheckEventArgs e) {
        if (e.NewValue == CheckState.Checked) {
            string taskName = taskList.Items[e.Index].ToString();
            this.BeginInvoke(new Action(delegate {
                if (taskDates.ContainsKey(taskName)) {
                    File.AppendAllText(completedLogFile, string.Format("[完成時間：{0}] 任務：{1} (建立於：{2})\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), taskName, taskDates[taskName].ToString("yyyy-MM-dd HH:mm:ss")));
                    taskDates.Remove(taskName);
                }
                taskList.Items.RemoveAt(e.Index); SaveActiveTasks();
            }));
        }
    }

    private void OnTaskDoubleClicked(object sender, MouseEventArgs e) {
        int index = taskList.IndexFromPoint(e.Location);
        if (index != ListBox.NoMatches) {
            string oldTask = taskList.Items[index].ToString();
            string newTask = ShowInputBox("請輸入修改後的任務內容：", "✏️ 修改待辦事項", oldTask);
            if (!string.IsNullOrWhiteSpace(newTask) && newTask != oldTask) {
                if (taskDates.ContainsKey(newTask)) { MessageBox.Show("名稱已存在！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                taskDates[newTask] = taskDates[oldTask]; taskDates.Remove(oldTask);
                taskList.Items[index] = newTask; SaveActiveTasks();
            }
        }
    }

    // 【全新開放的 API】：讓週期任務可以自動把項目塞進來
    public void AddTaskExternally(string text) {
        if (this.InvokeRequired) { this.BeginInvoke(new Action<string>(AddTaskExternally), new object[] { text }); return; }
        if (string.IsNullOrWhiteSpace(text)) return;
        if (taskDates.ContainsKey(text)) return; // 如果清單上已經有了(還沒做完)，就不重複加入
        
        DateTime now = DateTime.Now;
        taskList.Items.Insert(0, text); 
        taskDates[text] = now;

        File.AppendAllText(addedLogFile, string.Format("[{0}] 系統排程自動新增：{1}\r\n", now.ToString("yyyy-MM-dd HH:mm:ss"), text));
        SaveActiveTasks();
    }

    private static string ShowInputBox(string prompt, string title, string defaultValue) {
        Form form = new Form() { Width = 380, Height = 175, FormBorderStyle = FormBorderStyle.FixedDialog, Text = title, StartPosition = FormStartPosition.CenterScreen, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White };
        Label textLabel = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true, Font = MainFont };
        TextBox textBox = new TextBox() { Left = 20, Top = 60, Width = 320, Text = defaultValue, Font = MainFont };
        Button confirmation = new Button() { Text = "儲存修改", Left = 160, Width = 85, Top = 95, DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White }; confirmation.FlatAppearance.BorderSize = 0;
        Button cancel = new Button() { Text = "取消", Left = 255, Width = 85, Top = 95, DialogResult = DialogResult.Cancel, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(230,230,230) }; cancel.FlatAppearance.BorderSize = 0;
        confirmation.Click += new EventHandler(delegate { form.Close(); }); cancel.Click += new EventHandler(delegate { form.Close(); });
        form.Controls.Add(textBox); form.Controls.Add(confirmation); form.Controls.Add(cancel); form.Controls.Add(textLabel);
        form.AcceptButton = confirmation; form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : "";
    }
}
