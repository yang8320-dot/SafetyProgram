/*
 * 檔案功能：待辦事項與計畫大綱管理模組 (支援新增、刪除、以及跨清單轉移任務)
 * 對應選單名稱：待辦事項 / 計畫大綱 (動態生成)
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表名稱：TodoList
 */

using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_TodoList : UserControl
{
    private MainForm parentForm;
    public dynamic TargetList { get; set; } 

    private string listName;
    private string transferBtnText;
    private List<TodoItem> taskList = new List<TodoItem>();

    // --- 介面控制項 ---
    private Panel topInputPanel;
    private TextBox txtInput;
    private Button btnAdd;
    private ListBox listBoxTasks;
    private Panel bottomActionPanel;
    private Button btnTransfer;
    private Button btnDelete;

    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font DateFont = new Font("Microsoft JhengHei UI", 8.5f, FontStyle.Regular);

    // --- 資料模型 ---
    private class TodoItem
    {
        public string Id { get; set; }
        public string Content { get; set; }
        public string CreatedDate { get; set; }
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
        
        // 啟動時非同步載入 SQLite 資料
        _ = LoadDataAsync();
    }

    private void InitializeUI()
    {
        // 頂部輸入區塊
        topInputPanel = new Panel() { Dock = DockStyle.Top, Height = 45, BackColor = Color.White, Padding = new Padding(10), Margin = new Padding(0, 0, 0, 15) };

        btnAdd = new Button() { Text = "新增", Dock = DockStyle.Right, Width = 80, FlatStyle = FlatStyle.Flat, BackColor = AppleBlue, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold) };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += async (s, e) => await AddTaskAsync();

        txtInput = new TextBox() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, Font = new Font("Microsoft JhengHei UI", 12f, FontStyle.Regular), Margin = new Padding(0, 5, 10, 0) };
        txtInput.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; await AddTaskAsync(); } };

        topInputPanel.Controls.Add(txtInput);
        topInputPanel.Controls.Add(btnAdd);

        // 底部操作區塊
        bottomActionPanel = new Panel() { Dock = DockStyle.Bottom, Height = 45, Padding = new Padding(0, 15, 0, 0) };

        btnDelete = new Button() { Text = "刪除任務", Dock = DockStyle.Right, Width = 100, FlatStyle = FlatStyle.Flat, BackColor = AppleRed, ForeColor = Color.White, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular) };
        btnDelete.FlatAppearance.BorderSize = 0;
        btnDelete.Click += async (s, e) => await DeleteSelectedTaskAsync();

        btnTransfer = new Button() { Text = transferBtnText, Dock = DockStyle.Left, Width = 100, FlatStyle = FlatStyle.Flat, BackColor = Color.White, ForeColor = AppleBlue, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold) };
        btnTransfer.FlatAppearance.BorderSize = 1;
        btnTransfer.FlatAppearance.BorderColor = AppleBlue;
        btnTransfer.Click += async (s, e) => await TransferSelectedTaskAsync();

        bottomActionPanel.Controls.Add(btnDelete);
        bottomActionPanel.Controls.Add(btnTransfer);

        // 中間列表區塊
        listBoxTasks = new ListBox() { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White, IntegralHeight = false, ItemHeight = 60, DrawMode = DrawMode.OwnerDrawFixed };
        listBoxTasks.DrawItem += ListBoxTasks_DrawItem;

        this.Controls.Add(listBoxTasks);
        this.Controls.Add(topInputPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = AppleBgColor }); 
        this.Controls.Add(bottomActionPanel);
    }

    private void ListBoxTasks_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= taskList.Count) return;
        TodoItem item = taskList[e.Index];
        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

        Color bgColor = isSelected ? Color.FromArgb(235, 235, 240) : Color.White;
        using (SolidBrush bgBrush = new SolidBrush(bgColor)) { e.Graphics.FillRectangle(bgBrush, e.Bounds); }
        using (SolidBrush textBrush = new SolidBrush(Color.Black)) { e.Graphics.DrawString(item.Content, MainFont, textBrush, new Rectangle(e.Bounds.X + 15, e.Bounds.Y + 10, e.Bounds.Width - 30, 20)); }
        using (SolidBrush dateBrush = new SolidBrush(Color.Gray)) { e.Graphics.DrawString($"建立於: {item.CreatedDate}", DateFont, dateBrush, new Rectangle(e.Bounds.X + 15, e.Bounds.Y + 35, e.Bounds.Width - 30, 20)); }
        using (Pen linePen = new Pen(Color.FromArgb(230, 230, 230))) { e.Graphics.DrawLine(linePen, e.Bounds.X + 15, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1); }
        e.DrawFocusRectangle();
    }

    // ==========================================
    // SQLite 資料庫操作 (Async & Thread-Safety)
    // ==========================================

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
                    using (var cmd = new SQLiteCommand("SELECT Id, Content, CreatedDate FROM TodoList WHERE ListName = @ListName ORDER BY CreatedDate ASC", conn))
                    {
                        cmd.Parameters.AddWithValue("@ListName", this.listName);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                list.Add(new TodoItem
                                {
                                    Id = reader.GetString(0),
                                    Content = reader.GetString(1),
                                    CreatedDate = reader.GetString(2)
                                });
                            }
                        }
                    }
                }
                return list;
            });

            UpdateUIList(loadedTasks);
        }
        catch (Exception ex) { MessageBox.Show($"載入資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    private void UpdateUIList(List<TodoItem> newTasks = null)
    {
        if (this.InvokeRequired) { this.Invoke(new Action(() => UpdateUIList(newTasks))); return; }
        if (newTasks != null) taskList = newTasks;
        
        listBoxTasks.Items.Clear();
        foreach (var task in taskList) listBoxTasks.Items.Add(task);
    }

    private async Task AddTaskAsync()
    {
        string content = txtInput.Text.Trim();
        if (string.IsNullOrEmpty(content)) return;

        TodoItem newItem = new TodoItem
        {
            Id = Guid.NewGuid().ToString("N"),
            Content = content,
            CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") // 強制轉換一致時間格式
        };

        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("INSERT INTO TodoList (Id, ListName, Content, CreatedDate) VALUES (@Id, @ListName, @Content, @CreatedDate)", conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", newItem.Id);
                        cmd.Parameters.AddWithValue("@ListName", this.listName);
                        cmd.Parameters.AddWithValue("@Content", newItem.Content);
                        cmd.Parameters.AddWithValue("@CreatedDate", newItem.CreatedDate);
                        cmd.ExecuteNonQuery();
                    }
                }
            });

            taskList.Add(newItem);
            txtInput.Text = "";
            UpdateUIList();
        }
        catch (Exception ex) { MessageBox.Show($"新增失敗: {ex.Message}"); }
    }

    private async Task DeleteSelectedTaskAsync()
    {
        int selectedIndex = listBoxTasks.SelectedIndex;
        if (selectedIndex < 0 || selectedIndex >= taskList.Count) return;

        string targetId = taskList[selectedIndex].Id;

        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("DELETE FROM TodoList WHERE Id = @Id", conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", targetId);
                        cmd.ExecuteNonQuery();
                    }
                }
            });

            taskList.RemoveAt(selectedIndex);
            UpdateUIList();
        }
        catch (Exception ex) { MessageBox.Show($"刪除失敗: {ex.Message}"); }
    }

    private async Task TransferSelectedTaskAsync()
    {
        int selectedIndex = listBoxTasks.SelectedIndex;
        if (selectedIndex < 0 || selectedIndex >= taskList.Count) return;

        if (TargetList != null)
        {
            TodoItem itemToTransfer = taskList[selectedIndex];
            
            // 呼叫目標清單加入資料庫，然後當前清單刪除資料庫紀錄
            await TargetList.ReceiveTaskAsync(itemToTransfer.Content, itemToTransfer.CreatedDate);
            await DeleteTaskByIdAsync(itemToTransfer.Id); // 專屬的底層刪除方法

            taskList.RemoveAt(selectedIndex);
            UpdateUIList();
        }
    }

    /// <summary>
    /// 供內部呼叫：透過 ID 刪除 SQLite 紀錄 (不觸發 UI 更新)
    /// </summary>
    private async Task DeleteTaskByIdAsync(string id)
    {
        await Task.Run(() =>
        {
            using (var conn = DatabaseManager.GetConnection())
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM TodoList WHERE Id = @Id", conn))
                {
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        });
    }

    /// <summary>
    /// 供外部清單呼叫：接收轉移過來的任務並寫入 SQLite
    /// </summary>
    public async Task ReceiveTaskAsync(string content, string originalDate)
    {
        TodoItem newItem = new TodoItem
        {
            Id = Guid.NewGuid().ToString("N"),
            Content = content,
            CreatedDate = originalDate 
        };

        await Task.Run(() =>
        {
            using (var conn = DatabaseManager.GetConnection())
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("INSERT INTO TodoList (Id, ListName, Content, CreatedDate) VALUES (@Id, @ListName, @Content, @CreatedDate)", conn))
                {
                    cmd.Parameters.AddWithValue("@Id", newItem.Id);
                    cmd.Parameters.AddWithValue("@ListName", this.listName);
                    cmd.Parameters.AddWithValue("@Content", newItem.Content);
                    cmd.Parameters.AddWithValue("@CreatedDate", newItem.CreatedDate);
                    cmd.ExecuteNonQuery();
                }
            }
        });

        taskList.Add(newItem);
        UpdateUIList();
    }
}
