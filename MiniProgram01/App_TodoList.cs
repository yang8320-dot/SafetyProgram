using System;
using System.Drawing;
using System.Windows.Forms;

public class App_TodoList : UserControl {
    private MainForm parentForm;
    private TextBox inputField;
    private CheckedListBox taskList;
    private Button btnAdd;

    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);

    public App_TodoList(MainForm mainForm) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // --- 頂部輸入區 ---
        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 40 };
        
        inputField = new TextBox() { 
            Location = new Point(5, 5), 
            Width = 240, 
            Font = MainFont,
            BorderStyle = BorderStyle.FixedSingle
        };
        // 綁定 Enter 鍵快速新增
        inputField.KeyDown += new KeyEventHandler(OnInputKeyDown);

        btnAdd = new Button() { 
            Text = "新增", 
            Location = new Point(255, 4), 
            Width = 65, 
            Height = 26, 
            FlatStyle = FlatStyle.Flat, 
            BackColor = AppleBlue, 
            ForeColor = Color.White, 
            Font = new Font(MainFont.FontFamily, 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += new EventHandler(delegate { AddTask(); });

        topPanel.Controls.Add(inputField);
        topPanel.Controls.Add(btnAdd);
        this.Controls.Add(topPanel);

        // --- 待辦清單顯示區 ---
        taskList = new CheckedListBox() {
            Dock = DockStyle.Fill,
            Font = new Font(MainFont.FontFamily, 10.5f),
            CheckOnClick = true,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White
        };
        // 當打勾時，給予視覺回饋 (未來可擴充為刪除線或移動到完成區)
        taskList.ItemCheck += new ItemCheckEventHandler(OnTaskChecked);
        
        this.Controls.Add(taskList);
        taskList.BringToFront(); // 確保清單不會被頂部面板蓋住
        
        // 預設示範資料
        taskList.Items.Add("歡迎使用待辦事項模組！");
        taskList.Items.Add("這是一個獨立的頁籤架構");
    }

    private void AddTask() {
        string text = inputField.Text.Trim();
        if (!string.IsNullOrEmpty(text)) {
            taskList.Items.Insert(0, text); // 最新任務置頂
            inputField.Text = "";
            inputField.Focus();
        }
    }

    private void OnInputKeyDown(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) {
            e.SuppressKeyPress = true; // 消除叮叮聲
            AddTask();
        }
    }

    private void OnTaskChecked(object sender, ItemCheckEventArgs e) {
        // 未來這裡可以寫入「打勾後刪除」或是「打勾後存檔」的進階邏輯
    }
}
