/*
 * 檔案功能：待辦事項與計畫大綱管理模組 (支援新增、刪除、以及跨清單轉移任務)
 * 對應選單名稱：待辦事項 / 計畫大綱 (動態生成)
 * 對應資料庫名稱：(本模組採用純文字檔存儲) MainDB_TodoList_{listName}.txt
 * 資料表名稱：無 (資料欄位採用 '|' 符號間隔)
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_TodoList : UserControl
{
    // --- 子模組參考 ---
    private MainForm parentForm;
    public dynamic TargetList { get; set; } // 供外部設定目標清單，用於轉移任務

    // --- 核心變數 ---
    private string listName;
    private string transferBtnText;
    private string dataFilePath;
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

        public override string ToString()
        {
            // 寫入文字檔的格式，使用 '|' 間隔
            return $"{Id}|{Content}|{CreatedDate}";
        }
    }

    public App_TodoList(MainForm parent, string listName, string transferBtnText)
    {
        this.parentForm = parent;
        this.listName = listName;
        this.transferBtnText = transferBtnText;
        this.dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"MainDB_TodoList_{this.listName}.txt");

        // 1. 初始化控制項與 DPI 支援
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15); // 與主框架間保持 15 的留白

        // 2. 建構純程式碼 UI
        InitializeUI();

        // 3. 載入資料 (非同步)
        LoadDataAsync();
    }

    /// <summary>
    /// 建構 iOS 風格純程式碼介面 (Code-First UI)
    /// </summary>
    private void InitializeUI()
    {
        // ==========================================
        // 頂部輸入區塊 (Input Panel)
        // ==========================================
        topInputPanel = new Panel()
        {
            Dock = DockStyle.Top,
            Height = 45,
            BackColor = Color.White,
            Padding = new Padding(10), // 框內與文字間隔 10
            Margin = new Padding(0, 0, 0, 15) // 與下方列表間距 15
        };

        btnAdd = new Button()
        {
            Text = "新增",
            Dock = DockStyle.Right,
            Width = 80,
            FlatStyle = FlatStyle.Flat,
            BackColor = AppleBlue,
            ForeColor = Color.White,
            Cursor = Cursors.Hand,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold)
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += async (s, e) => await AddTaskAsync();

        txtInput = new TextBox()
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            Font = new Font("Microsoft JhengHei UI", 12f, FontStyle.Regular),
        };
        // 處理 TextBox 在 Panel 中的垂直置中
        txtInput.Margin = new Padding(0, 5, 10, 0); 
        txtInput.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; await AddTaskAsync(); } };

        topInputPanel.Controls.Add(txtInput);
        topInputPanel.Controls.Add(btnAdd);

        // ==========================================
        // 底部操作區塊 (Action Panel)
        // ==========================================
        bottomActionPanel = new Panel()
        {
            Dock = DockStyle.Bottom,
            Height = 45,
            Padding = new Padding(0, 15, 0, 0) // 與上方列表間距 15
        };

        btnDelete = new Button()
        {
            Text = "刪除任務",
            Dock = DockStyle.Right,
            Width = 100,
            FlatStyle = FlatStyle.Flat,
            BackColor = AppleRed,
            ForeColor = Color.White,
            Cursor = Cursors.Hand,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular)
        };
        btnDelete.FlatAppearance.BorderSize = 0;
        btnDelete.Click += async (s, e) => await DeleteSelectedTaskAsync();

        btnTransfer = new Button()
        {
            Text = transferBtnText,
            Dock = DockStyle.Left,
            Width = 100,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            ForeColor = AppleBlue,
            Cursor = Cursors.Hand,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold)
        };
        btnTransfer.FlatAppearance.BorderSize = 1;
        btnTransfer.FlatAppearance.BorderColor = AppleBlue;
        btnTransfer.Click += async (s, e) => await TransferSelectedTaskAsync();

        bottomActionPanel.Controls.Add(btnDelete);
        bottomActionPanel.Controls.Add(btnTransfer);

        // ==========================================
        // 中間列表區塊 (iOS Style ListBox)
        // ==========================================
        listBoxTasks = new ListBox()
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            IntegralHeight = false,
            ItemHeight = 60, // 加大行高，適應觸控與閱讀
            DrawMode = DrawMode.OwnerDrawFixed // 啟用自訂繪製
        };
        listBoxTasks.DrawItem += ListBoxTasks_DrawItem;

        // ==========================================
        // 組合至主畫面 (注意加入順序會影響 Dock 排版)
        // ==========================================
        this.Controls.Add(listBoxTasks);
        this.Controls.Add(topInputPanel);
        // 加入分隔用透明 Panel 實作 Margin 效果
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = AppleBgColor }); 
        this.Controls.Add(bottomActionPanel);
    }

    /// <summary>
    /// 自訂繪製 ListBox 項目 (iOS 儲存格風格)
    /// </summary>
    private void ListBoxTasks_DrawItem(object sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= taskList.Count) return;

        TodoItem item = taskList[e.Index];
        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

        // 設定背景色
        Color bgColor = isSelected ? Color.FromArgb(235, 235, 240) : Color.White;
        using (SolidBrush bgBrush = new SolidBrush(bgColor))
        {
            e.Graphics.FillRectangle(bgBrush, e.Bounds);
        }

        // 繪製文字 (標題)
        using (SolidBrush textBrush = new SolidBrush(Color.Black))
        {
            Rectangle textRect = new Rectangle(e.Bounds.X + 15, e.Bounds.Y + 10, e.Bounds.Width - 30, 20);
            e.Graphics.DrawString(item.Content, MainFont, textBrush, textRect);
        }

        // 繪製文字 (日期)
        using (SolidBrush dateBrush = new SolidBrush(Color.Gray))
        {
            Rectangle dateRect = new Rectangle(e.Bounds.X + 15, e.Bounds.Y + 35, e.Bounds.Width - 30, 20);
            e.Graphics.DrawString($"建立於: {item.CreatedDate}", DateFont, dateBrush, dateRect);
        }

        // 繪製底部灰色分隔線 (iOS Style)
        using (Pen linePen = new Pen(Color.FromArgb(230, 230, 230)))
        {
            e.Graphics.DrawLine(linePen, e.Bounds.X + 15, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        e.DrawFocusRectangle();
    }

    // ==========================================
    // 核心邏輯與非同步檔案處理 (Thread-Safety)
    // ==========================================

    /// <summary>
    /// 非同步載入資料
    /// </summary>
    private async Task LoadDataAsync()
    {
        if (!File.Exists(dataFilePath)) return;

        try
        {
            string[] lines = await Task.Run(() => File.ReadAllLines(dataFilePath));
            List<TodoItem> loadedTasks = new List<TodoItem>();

            foreach (var line in lines)
            {
                string[] parts = line.Split('|');
                if (parts.Length >= 3)
                {
                    loadedTasks.Add(new TodoItem
                    {
                        Id = parts[0],
                        Content = parts[1],
                        CreatedDate = parts[2]
                    });
                }
            }

            // 安全更新 UI
            UpdateUIList(loadedTasks);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"讀取資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// 非同步儲存資料
    /// </summary>
    private async Task SaveDataAsync()
    {
        try
        {
            List<string> lines = taskList.Select(t => t.ToString()).ToList();
            await Task.Run(() => File.WriteAllLines(dataFilePath, lines));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"儲存資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// 更新 UI 列表 (實作 InvokeRequired 確保 Thread-Safety)
    /// </summary>
    private void UpdateUIList(List<TodoItem> newTasks = null)
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(() => UpdateUIList(newTasks)));
            return;
        }

        if (newTasks != null)
        {
            taskList = newTasks;
        }

        listBoxTasks.Items.Clear();
        foreach (var task in taskList)
        {
            listBoxTasks.Items.Add(task); // 加入物件以配合 OwnerDrawFixed
        }
    }

    /// <summary>
    /// 新增任務
    /// </summary>
    private async Task AddTaskAsync()
    {
        string content = txtInput.Text.Trim();
        if (string.IsNullOrEmpty(content)) return;

        TodoItem newItem = new TodoItem
        {
            Id = Guid.NewGuid().ToString("N"),
            Content = content,
            CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") // 強制統一日期格式
        };

        taskList.Add(newItem);
        txtInput.Text = "";
        
        UpdateUIList();
        await SaveDataAsync();
    }

    /// <summary>
    /// 刪除選取任務
    /// </summary>
    private async Task DeleteSelectedTaskAsync()
    {
        int selectedIndex = listBoxTasks.SelectedIndex;
        if (selectedIndex < 0 || selectedIndex >= taskList.Count) return;

        taskList.RemoveAt(selectedIndex);
        UpdateUIList();
        await SaveDataAsync();
    }

    /// <summary>
    /// 跨清單轉移任務
    /// </summary>
    private async Task TransferSelectedTaskAsync()
    {
        int selectedIndex = listBoxTasks.SelectedIndex;
        if (selectedIndex < 0 || selectedIndex >= taskList.Count) return;

        if (TargetList != null)
        {
            TodoItem itemToTransfer = taskList[selectedIndex];
            
            // 呼叫目標清單的公開方法加入任務
            await TargetList.ReceiveTaskAsync(itemToTransfer.Content, itemToTransfer.CreatedDate);

            // 從當前清單移除
            taskList.RemoveAt(selectedIndex);
            UpdateUIList();
            await SaveDataAsync();
        }
    }

    /// <summary>
    /// 供外部清單呼叫，接收轉移過來的任務
    /// </summary>
    public async Task ReceiveTaskAsync(string content, string originalDate)
    {
        TodoItem newItem = new TodoItem
        {
            Id = Guid.NewGuid().ToString("N"),
            Content = content,
            CreatedDate = originalDate // 保留原始建立時間
        };

        taskList.Add(newItem);
        UpdateUIList();
        await SaveDataAsync();
    }
}
