using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_TodoList : UserControl {
    private string activeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_active.txt");
    private string addedLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_added.txt");
    private string doneLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "todo_history_completed.txt");
    private TextBox inputField;
    private CheckedListBox taskList;
    private Dictionary<string, DateTime> taskDates = new Dictionary<string, DateTime>();
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);

    public App_TodoList(MainForm parent) {
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        Panel top = new Panel() { Dock = DockStyle.Top, Height = 35 };
        inputField = new TextBox() { Width = 240, Font = new Font("Microsoft JhengHei UI", 10f) };
        inputField.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AddTask(inputField.Text); } };
        Button btnAdd = new Button() { Text = "新增", Left = 250, Width = 65, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White };
        btnAdd.Click += (s, e) => AddTask(inputField.Text);
        top.Controls.AddRange(new Control[] { inputField, btnAdd });

        taskList = new CheckedListBox() { Dock = DockStyle.Fill, CheckOnClick = true, BorderStyle = BorderStyle.None, Font = new Font("Microsoft JhengHei UI", 10.5f) };
        taskList.ItemCheck += OnTaskChecked;
        taskList.MouseDoubleClick += OnTaskEdit;

        this.Controls.Add(taskList); this.Controls.Add(top);
        LoadTasks();
    }

    public void AddTaskExternally(string text) { if (!taskDates.ContainsKey(text)) AddTask(text, true); }

    private void AddTask(string text, bool auto = false) {
        text = text.Trim(); if (string.IsNullOrEmpty(text) || taskDates.ContainsKey(text)) return;
        DateTime now = DateTime.Now;
        taskList.Items.Insert(0, text); taskDates[text] = now;
        File.AppendAllText(addedLog, string.Format("[{0}] {1}: {2}\n", now.ToString("yyyy-MM-dd HH:mm"), auto ? "排程" : "手動", text));
        SaveActive(); inputField.Text = "";
    }

    private void OnTaskChecked(object sender, ItemCheckEventArgs e) {
        if (e.NewValue == CheckState.Checked) {
            string name = taskList.Items[e.Index].ToString();
            this.BeginInvoke(new Action(() => {
                if (taskDates.ContainsKey(name)) {
                    File.AppendAllText(doneLog, string.Format("[完成:{0}] {1} (建立於:{2})\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), name, taskDates[name].ToString("yyyy-MM-dd HH:mm")));
                    taskDates.Remove(name);
                }
                taskList.Items.RemoveAt(e.Index); SaveActive();
            }));
        }
    }

    private void OnTaskEdit(object sender, MouseEventArgs e) {
        int idx = taskList.IndexFromPoint(e.Location);
        if (idx != -1) {
            string old = taskList.Items[idx].ToString();
            string nw = App_Shortcuts.ShowInputBox("修改任務：", "✏️ 編輯", old);
            if (!string.IsNullOrEmpty(nw) && nw != old && !taskDates.ContainsKey(nw)) {
                taskDates[nw] = taskDates[old]; taskDates.Remove(old);
                taskList.Items[idx] = nw; SaveActive();
            }
        }
    }

    private void SaveActive() {
        List<string> lines = new List<string>();
        foreach (var item in taskList.Items) { string n = item.ToString(); lines.Add(n + "|" + taskDates[n].ToString()); }
        File.WriteAllLines(activeFile, lines);
    }

    private void LoadTasks() {
        if (!File.Exists(activeFile)) return;
        foreach (string l in File.ReadAllLines(activeFile)) {
            string[] p = l.Split('|');
            if (p.Length >= 2) { taskList.Items.Add(p[0]); taskDates[p[0]] = DateTime.Parse(p[1]); }
        }
    }
}
