/*
 * 檔案功能：待辦事項與計畫大綱管理模組 (Microsoft.Data.Sqlite 升級版)
 */

using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_TodoList : UserControl
{
    private MainForm parentForm;
    public dynamic TargetList { get; set; } 

    private string listName;
    private string transferBtnText;
    private List<TodoItem> taskList = new List<TodoItem>();

    private Panel topInputPanel;
    private TextBox txtInput;
    private Button btnAdd;
    private FlowLayoutPanel taskPanel;

    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleOrange = Color.FromArgb(255, 149, 0);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);
    private static Font SmallFont = new Font("Microsoft JhengHei UI", 9.5f, FontStyle.Regular);

    private readonly string[] colorCycle = { "Black", "AppleRed", "AppleBlue", "ApplePurple", "AppleGreen", "AppleOrange" };

    private Color GetDisplayColor(string colorName)
    {
        switch (colorName)
        {
            case "AppleRed": return Color.FromArgb(255, 59, 48);
            case "AppleBlue": return Color.FromArgb(0, 122, 255);
            case "ApplePurple": return Color.FromArgb(175, 82, 222);
            case "AppleGreen": return Color.FromArgb(52, 199, 89);
            case "AppleOrange": return Color.FromArgb(255, 149, 0);
            default: return Color.Black;
        }
    }

    public class TodoItem
    {
        public string Id { get; set; }
        public string Content { get; set; }
        public string CreatedDate { get; set; }
        public string Color { get; set; }
        public string Note { get; set; }
    }

    public App_TodoList(MainForm parent, string listName, string transferBtnText)
    {
        this.parentForm = parent;
        this.listName = listName;
        this.transferBtnText = transferBtnText;

        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15); 

        InitializeUI();
        _ = LoadDataAsync();
    }

    private void InitializeUI()
    {
        topInputPanel = new Panel() { Dock = DockStyle.Top, Height = 45, BackColor = Color.White, Padding = new Padding(10), Margin = new Padding(0, 0, 0, 15) };
        btnAdd = new Button() { Text = "新增", Dock = DockStyle.Right, Width = 80, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold) };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += async (s, e) => await AddTaskAsync();

        txtInput = new TextBox() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, Font = new Font("Microsoft JhengHei UI", 12f, FontStyle.Regular), Margin = new Padding(0, 5, 10, 0) };
        txtInput.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; await AddTaskAsync(); } };

        topInputPanel.Controls.Add(txtInput);
        topInputPanel.Controls.Add(btnAdd);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = AppleBgColor };
        taskPanel.Resize += (s, e) =>
        {
            int safeWidth = taskPanel.ClientSize.Width - 20;
            if (safeWidth > 0) { foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth; }
        };

        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = AppleBgColor }); 
        this.Controls.Add(topInputPanel);
    }

    public void RefreshUI()
    {
        if (this.InvokeRequired) { this.Invoke(new Action(RefreshUI)); return; }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var t in taskList)
        {
            Panel card = new Panel() { Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 65), Margin = new Padding(0, 0, 0, 15), BackColor = Color.White, Padding = new Padding(10) };
            
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, AutoSize = true };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 45f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); 

            Button btnTransfer = new Button() { Text = transferBtnText, Dock = DockStyle.Fill, BackColor = Color.White, ForeColor = AppleBlue, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold), Margin = new Padding(0, 0, 10, 0) };
            btnTransfer.FlatAppearance.BorderColor = AppleBlue;
            btnTransfer.Click += async (s, e) => await TransferTaskAsync(t);
            
            TableLayoutPanel textPanel = new TableLayoutPanel() { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, AutoSize = true, Margin = new Padding(0, 0, 10, 0) };
            Label lblTitle = new Label() { Text = t.Content, ForeColor = GetDisplayColor(t.Color), Font = BoldFont, AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            Label lblDate = new Label() { Text = $"建立於: {t.CreatedDate}", ForeColor = Color.Gray, Font = SmallFont, AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            textPanel.Controls.Add(lblTitle, 0, 0); textPanel.Controls.Add(lblDate, 0, 1);

            Button btnColor = new Button() { Text = "色", Dock = DockStyle.Fill, BackColor = GetDisplayColor(t.Color), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(0, 0, 5, 0) };
            btnColor.FlatAppearance.BorderSize = 0;
            btnColor.Click += async (s, e) => {
                int idx = Array.IndexOf(colorCycle, t.Color);
                if (idx == -1) idx = 0;
                t.Color = colorCycle[(idx + 1) % colorCycle.Length];
                await UpdateTaskFieldAsync(t.Id, "Color", t.Color);
                RefreshUI();
            };

            Button btnNote = new Button() { Text = "註", Dock = DockStyle.Fill, BackColor = string.IsNullOrEmpty(t.Note) ? Color.FromArgb(230, 230, 230) : AppleOrange, ForeColor = string.IsNullOrEmpty(t.Note) ? Color.Gray : Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = SmallFont, Margin = new Padding(0, 0, 10, 0) };
            btnNote.FlatAppearance.BorderSize = 0;
            btnNote.Click += (s, e) => { new EditNoteWindow(this, t).ShowDialog(); };

            Button btnDel = new Button() { Text = "刪除", Dock = DockStyle.Fill, BackColor = AppleRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular), Margin = new Padding(0) };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += async (s, e) => {
                if (MessageBox.Show($"確定刪除任務【{t.Content}】？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) {
                    await DeleteTaskByIdAsync(t.Id);
                    taskList.Remove(t);
                    SafeAppendLog("刪除", t.Content);
                    RefreshUI();
                }
            };

            tlp.Controls.Add(btnTransfer, 0, 0); tlp.Controls.Add(textPanel, 1, 0);
            tlp.Controls.Add(btnColor, 2, 0); tlp.Controls.Add(btnNote, 3, 0); tlp.Controls.Add(btnDel, 4, 0);

            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }

    private void SafeAppendLog(string action, string content)
    {
        try
        {
            string logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MainDB_HistoryLog.txt");
            string dName = listName == "todo" ? "待辦" : "計畫";
            string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{dName}] {action}: {content}\r\n";
            File.AppendAllText(logFile, logLine);
        }
        catch { }
    }

    private async Task LoadDataAsync()
    {
        try
        {
            var loadedTasks = await Task.Run(() =>
            {
                var list = new List<TodoItem>();
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Id, Content, CreatedDate, Color, Note FROM TodoList WHERE ListName = @ListName ORDER BY CreatedDate ASC", conn))
                    {
                        cmd.Parameters.AddWithValue("@ListName", this.listName);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                list.Add(new TodoItem
                                {
                                    Id = reader.GetString(0), Content = reader.GetString(1), CreatedDate = reader.GetString(2),
                                    Color = reader.IsDBNull(3) ? "Black" : reader.GetString(3),
                                    Note = reader.IsDBNull(4) ? "" : reader.GetString(4)
                                });
                            }
                        }
                    }
                }
                return list;
            });
            taskList = loadedTasks;
            RefreshUI();
        }
        catch (Exception ex) { MessageBox.Show($"載入資料失敗: {ex.Message}"); }
    }

    private async Task AddTaskAsync()
    {
        string content = txtInput.Text.Trim();
        if (string.IsNullOrEmpty(content)) return;

        TodoItem newItem = new TodoItem { Id = Guid.NewGuid().ToString("N"), Content = content, CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), Color = "Black", Note = "" };

        try
        {
            await Task.Run(() => {
                using (var conn = DatabaseManager.GetConnection()) {
                    conn.Open();
                    using (var cmd = new SqliteCommand("INSERT INTO TodoList (Id, ListName, Content, CreatedDate, Color, Note) VALUES (@Id, @ListName, @Content, @CreatedDate, @Color, @Note)", conn)) {
                        cmd.Parameters.AddWithValue("@Id", newItem.Id); cmd.Parameters.AddWithValue("@ListName", this.listName);
                        cmd.Parameters.AddWithValue("@Content", newItem.Content); cmd.Parameters.AddWithValue("@CreatedDate", newItem.CreatedDate);
                        cmd.Parameters.AddWithValue("@Color", newItem.Color); cmd.Parameters.AddWithValue("@Note", newItem.Note);
                        cmd.ExecuteNonQuery();
                    }
                }
            });

            taskList.Add(newItem);
            SafeAppendLog("新增", newItem.Content);
            txtInput.Text = ""; RefreshUI();
        }
        catch (Exception ex) { MessageBox.Show($"新增失敗: {ex.Message}"); }
    }

    private async Task TransferTaskAsync(TodoItem t)
    {
        if (TargetList != null)
        {
            try {
                await TargetList.ReceiveTaskAsync(t.Content, t.CreatedDate, t.Color, t.Note);
                await DeleteTaskByIdAsync(t.Id);
                taskList.Remove(t);
                SafeAppendLog("轉移", t.Content);
                RefreshUI();
            } 
            catch (Exception ex) { MessageBox.Show("轉移失敗: " + ex.Message); }
        }
    }

    public async Task UpdateTaskFieldAsync(string id, string field, string value)
    {
        await Task.Run(() => {
            using (var conn = DatabaseManager.GetConnection()) {
                conn.Open();
                using (var cmd = new SqliteCommand($"UPDATE TodoList SET {field} = @Value WHERE Id = @Id", conn)) {
                    cmd.Parameters.AddWithValue("@Value", value); cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery();
                }
            }
        });
    }

    private async Task DeleteTaskByIdAsync(string id)
    {
        await Task.Run(() => {
            using (var conn = DatabaseManager.GetConnection()) {
                conn.Open();
                using (var cmd = new SqliteCommand("DELETE FROM TodoList WHERE Id = @Id", conn)) { cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery(); }
            }
        });
    }

    public async Task ReceiveTaskAsync(string content, string originalDate, string color = "Black", string note = "")
    {
        TodoItem newItem = new TodoItem { Id = Guid.NewGuid().ToString("N"), Content = content, CreatedDate = originalDate, Color = color, Note = note };

        await Task.Run(() => {
            using (var conn = DatabaseManager.GetConnection()) {
                conn.Open();
                using (var cmd = new SqliteCommand("INSERT INTO TodoList (Id, ListName, Content, CreatedDate, Color, Note) VALUES (@Id, @ListName, @Content, @CreatedDate, @Color, @Note)", conn)) {
                    cmd.Parameters.AddWithValue("@Id", newItem.Id); cmd.Parameters.AddWithValue("@ListName", this.listName);
                    cmd.Parameters.AddWithValue("@Content", newItem.Content); cmd.Parameters.AddWithValue("@CreatedDate", newItem.CreatedDate);
                    cmd.Parameters.AddWithValue("@Color", newItem.Color); cmd.Parameters.AddWithValue("@Note", newItem.Note);
                    cmd.ExecuteNonQuery();
                }
            }
        });

        taskList.Add(newItem); RefreshUI();
    }
}

public class EditNoteWindow : Form
{
    private App_TodoList parent;
    private App_TodoList.TodoItem item;
    private TextBox txtNote;

    public EditNoteWindow(App_TodoList parent, App_TodoList.TodoItem item)
    {
        this.parent = parent; this.item = item;
        
        this.Text = "任務備註";
        this.Width = 450; this.Height = 350; this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog; this.MaximizeBox = false; this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247); this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };
        f.Controls.Add(new Label() { Text = $"【 {item.Content} 】", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 122, 255), Margin = new Padding(0, 0, 0, 10) });

        txtNote = new TextBox() { Width = 390, Height = 180, Multiline = true, Text = item.Note ?? "", BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 15) };
        f.Controls.Add(txtNote);

        Button btnSave = new Button() { Text = "儲存備註", Width = 390, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) => {
            item.Note = txtNote.Text;
            await parent.UpdateTaskFieldAsync(item.Id, "Note", item.Note);
            parent.RefreshUI();
            this.Close();
        };

        f.Controls.Add(btnSave); this.Controls.Add(f);
    }
}
