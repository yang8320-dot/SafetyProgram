/*
 * 檔案功能：常用捷徑管理與開啟模組 (支援新增、編輯、刪除、拖曳排序與直接啟動程式/檔案/網頁)
 * 對應選單名稱：常用捷徑
 * 對應資料庫名稱：(本模組採用純文字檔存儲) MainDB_Shortcuts.txt
 * 資料表名稱：無 (資料欄位採用 '|' 符號間隔，路徑欄位經 Base64 保護)
 */

using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading.Tasks;

public class App_Shortcuts : UserControl
{
    private MainForm parentForm;
    private string shortcutFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MainDB_Shortcuts.txt");
    private FlowLayoutPanel taskPanel;

    // --- 拖曳排序相關變數 ---
    private int dragInsertIndex = -1;

    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleGreen = Color.FromArgb(52, 199, 89);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    // --- 資料模型 ---
    public class ShortcutItem
    {
        public string Name { get; set; }
        public string Path { get; set; }
    }

    public List<ShortcutItem> shortcuts = new List<ShortcutItem>();

    public App_Shortcuts(MainForm mainForm)
    {
        this.parentForm = mainForm;
        
        // 1. 初始化控制項與 DPI 支援
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15); // 與主框架間保持 15 的留白

        // 2. 建構純程式碼 UI
        InitializeUI();

        // 3. 載入資料 (非同步)
        _ = LoadShortcutsAsync();
    }

    /// <summary>
    /// 建構 iOS 風格純程式碼介面 (Code-First UI)
    /// </summary>
    private void InitializeUI()
    {
        // ==========================================
        // 頂部標題與控制區塊
        // ==========================================
        TableLayoutPanel header = new TableLayoutPanel()
        {
            Dock = DockStyle.Top,
            Height = 45,
            ColumnCount = 2,
            BackColor = Color.Transparent
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));

        Label lblTitle = new Label()
        {
            Text = "常用捷徑",
            Font = new Font(MainFont.FontFamily, 14f, FontStyle.Bold),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(5, 0, 0, 0)
        };

        Button btnAdd = new Button()
        {
            Text = "新增",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = AppleBlue,
            ForeColor = Color.White,
            Font = BoldFont,
            Margin = new Padding(0, 5, 0, 5) // 內縮按鈕以適應 Header 高度
        };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += (s, e) => { new EditShortcutWindow(this, -1, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);

        // ==========================================
        // 中間列表區塊 (支援拖曳排序)
        // ==========================================
        taskPanel = new FlowLayoutPanel()
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = AppleBgColor,
            AllowDrop = true
        };

        // 綁定拖曳相關事件
        taskPanel.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskPanel.DragOver += OnTaskDragOver;
        taskPanel.DragLeave += (s, e) => { dragInsertIndex = -1; taskPanel.Invalidate(); };
        taskPanel.DragDrop += OnTaskDragDrop; // 非同步事件
        taskPanel.Paint += OnTaskContainerPaint;

        // 確保內部卡片動態跟隨容器寬度
        taskPanel.Resize += (s, e) =>
        {
            int safeWidth = taskPanel.ClientSize.Width - 20; // 預留右側卷軸寬度
            if (safeWidth > 0)
            {
                foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
            }
        };

        // ==========================================
        // 組合至主畫面
        // ==========================================
        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent }); // Header 下方間距 15
        this.Controls.Add(header);
        
        taskPanel.BringToFront();
    }

    /// <summary>
    /// 更新介面顯示 (執行緒安全)
    /// </summary>
    public void RefreshUI()
    {
        if (this.InvokeRequired)
        {
            this.Invoke(new Action(RefreshUI));
            return;
        }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var s in shortcuts)
        {
            // 建立 iOS 風格懸浮卡片
            Panel card = new Panel()
            {
                Width = startWidth,
                AutoSize = true,
                MinimumSize = new Size(0, 55),
                Margin = new Padding(0, 0, 0, 15), // 卡片與卡片之間距離 15
                BackColor = Color.White,
                Padding = new Padding(10) // 框內與元件間隔 10
            };

            TableLayoutPanel tlp = new TableLayoutPanel()
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                AutoSize = true
            };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); // 啟動按鈕
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); // 刪除按鈕
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // 標題文字
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); // 編輯按鈕

            // 啟動按鈕
            Button btnOpen = new Button()
            {
                Text = "啟動",
                Dock = DockStyle.Top,
                Height = 35,
                BackColor = AppleGreen,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = BoldFont,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnOpen.FlatAppearance.BorderSize = 0;
            btnOpen.Click += (sender, e) =>
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo() { FileName = s.Path, UseShellExecute = true };
                    Process.Start(psi);
                }
                catch
                {
                    MessageBox.Show("無法開啟此捷徑，請檢查路徑或檔案是否存在！", "開啟失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // 刪除按鈕
            Button btnDel = new Button()
            {
                Text = "刪除",
                Dock = DockStyle.Top,
                Height = 35,
                BackColor = AppleRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = BoldFont,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += async (sender, e) =>
            {
                if (MessageBox.Show($"確定移除捷徑【{s.Name}】？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    shortcuts.Remove(s);
                    await SaveShortcutsAsync();
                    RefreshUI();
                }
            };

            // 標題文字 (可拖曳)
            Label lbl = new Label()
            {
                Text = s.Name,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
                Font = MainFont,
                Padding = new Padding(10, 0, 10, 0),
                Cursor = Cursors.SizeAll // 讓使用者知道可以拖曳
            };
            lbl.MouseDown += (sender, e) =>
            {
                if (e.Button == MouseButtons.Left) card.DoDragDrop(card, DragDropEffects.Move);
            };

            // 編輯按鈕
            Button btnEdit = new Button()
            {
                Text = "編輯",
                Dock = DockStyle.Top,
                Height = 35,
                BackColor = AppleBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = BoldFont,
                Margin = new Padding(10, 0, 0, 0)
            };
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (sender, e) =>
            {
                int idx = shortcuts.IndexOf(s);
                new EditShortcutWindow(this, idx, s).ShowDialog();
            };

            tlp.Controls.Add(btnOpen, 0, 0);
            tlp.Controls.Add(btnDel, 1, 0);
            tlp.Controls.Add(lbl, 2, 0);
            tlp.Controls.Add(btnEdit, 3, 0);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    // ==========================================
    // 拖曳排序核心邏輯
    // ==========================================
    private void OnTaskDragOver(object sender, DragEventArgs e)
    {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskPanel.PointToClient(new Point(e.X, e.Y));
        Control target = taskPanel.GetChildAtPoint(clientPoint);
        
        if (target != null && target is Panel)
        {
            int idx = taskPanel.Controls.GetChildIndex(target);
            // 判斷滑鼠在卡片的上半部還是下半部
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            
            if (dragInsertIndex != idx)
            {
                dragInsertIndex = idx;
                taskPanel.Invalidate(); // 觸發重新繪製插入提示線
            }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e)
    {
        if (dragInsertIndex != -1 && taskPanel.Controls.Count > 0)
        {
            // 繪製藍色分隔線
            int y = (dragInsertIndex < taskPanel.Controls.Count) ?
                taskPanel.Controls[dragInsertIndex].Top - 7 : 
                taskPanel.Controls[taskPanel.Controls.Count - 1].Bottom + 7;
            
            e.Graphics.FillRectangle(new SolidBrush(AppleBlue), 5, y, taskPanel.Width - 10, 3);
        }
    }

    private async void OnTaskDragDrop(object sender, DragEventArgs e)
    {
        Panel draggedCard = e.Data.GetData(typeof(Panel)) as Panel;
        if (draggedCard != null && dragInsertIndex != -1)
        {
            int targetIdx = dragInsertIndex;
            int currentIdx = taskPanel.Controls.GetChildIndex(draggedCard);
            
            if (currentIdx < targetIdx) targetIdx--; 
            
            // 同步記憶體內的 List 順序
            var item = shortcuts[currentIdx];
            shortcuts.RemoveAt(currentIdx);
            
            int finalIdx = Math.Min(targetIdx, shortcuts.Count);
            shortcuts.Insert(finalIdx, item);

            dragInsertIndex = -1; 
            taskPanel.Invalidate(); 
            
            await SaveShortcutsAsync(); // 非同步儲存新順序
            RefreshUI();
        }
    }

    // ==========================================
    // 存檔與載入 (非同步 I/O Thread-Safety)
    // ==========================================
    public async Task SaveShortcutsAsync()
    {
        try
        {
            List<string> lines = new List<string>();
            foreach(var s in shortcuts)
            {
                // 將路徑進行 Base64 編碼，避免路徑字串中含有 '|' 破壞資料結構
                string base64Path = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(s.Path));
                lines.Add($"{s.Name}|{base64Path}");
            }
            await Task.Run(() => File.WriteAllLines(shortcutFile, lines));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"儲存捷徑失敗: {ex.Message}");
        }
    }

    public async Task LoadShortcutsAsync()
    {
        if(!File.Exists(shortcutFile)) return;
        
        try
        {
            string[] lines = await Task.Run(() => File.ReadAllLines(shortcutFile));
            shortcuts.Clear();

            foreach(var l in lines)
            {
                var p = l.Split('|');
                if(p.Length >= 2)
                {
                    try
                    {
                        // 嘗試 Base64 解碼
                        string path = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(p[1]));
                        shortcuts.Add(new ShortcutItem() { Name = p[0], Path = path });
                    }
                    catch
                    {
                        // 相容舊版未加密的資料
                        shortcuts.Add(new ShortcutItem() { Name = p[0], Path = p[1] });
                    }
                }
            }
            RefreshUI();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"載入捷徑失敗: {ex.Message}");
        }
    }
}

// ==========================================
// 視窗：新增/編輯捷徑 (iOS 風格)
// ==========================================
public class EditShortcutWindow : Form
{
    private App_Shortcuts parent;
    private int index;
    private TextBox txtName, txtPath;

    public EditShortcutWindow(App_Shortcuts p, int idx, App_Shortcuts.ShortcutItem item)
    {
        this.parent = p;
        this.index = idx;
        
        // 視窗基本設定
        this.Text = idx == -1 ? "新增捷徑" : "編輯捷徑";
        this.Width = 450; 
        this.Height = 300;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false; 
        this.MinimizeBox = false;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        // 主容器
        FlowLayoutPanel f = new FlowLayoutPanel()
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            Padding = new Padding(20)
        };

        // 欄位 1：名稱
        f.Controls.Add(new Label() { Text = "捷徑名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        txtName = new TextBox()
        {
            Width = 380,
            Text = item?.Name ?? "",
            Margin = new Padding(0, 0, 0, 15),
            BorderStyle = BorderStyle.FixedSingle
        };
        f.Controls.Add(txtName);

        // 欄位 2：路徑
        f.Controls.Add(new Label() { Text = "目標路徑 (檔案 / 資料夾 / 網址)：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        
        TableLayoutPanel pathRow = new TableLayoutPanel()
        {
            Width = 380,
            Height = 35,
            ColumnCount = 2,
            Margin = new Padding(0, 0, 0, 20)
        };
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));
        
        txtPath = new TextBox()
        {
            Dock = DockStyle.Fill,
            Text = item?.Path ?? "",
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 5, 10, 0) // 垂直置中微調
        };

        Button btnBrowse = new Button()
        {
            Text = "瀏覽",
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = Color.White
        };
        btnBrowse.FlatAppearance.BorderColor = Color.Gray;
        btnBrowse.Click += (s, e) =>
        {
            OpenFileDialog ofd = new OpenFileDialog() { Title = "選擇捷徑目標檔案" };
            if (ofd.ShowDialog() == DialogResult.OK) { txtPath.Text = ofd.FileName; }
        };

        pathRow.Controls.Add(txtPath, 0, 0);
        pathRow.Controls.Add(btnBrowse, 1, 0);
        f.Controls.Add(pathRow);

        // 儲存按鈕
        Button btnSave = new Button()
        {
            Text = "儲存設定",
            Width = 380,
            Height = 40,
            BackColor = Color.FromArgb(0, 122, 255), // Apple Blue
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold)
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtPath.Text))
            {
                MessageBox.Show("名稱與路徑不可為空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (index == -1)
            {
                parent.shortcuts.Add(new App_Shortcuts.ShortcutItem() { Name = txtName.Text, Path = txtPath.Text });
            }
            else
            {
                parent.shortcuts[index].Name = txtName.Text;
                parent.shortcuts[index].Path = txtPath.Text;
            }

            // 非同步等待主視窗存檔完畢後再關閉
            await parent.SaveShortcutsAsync();
            parent.RefreshUI();
            this.Close();
        };
        
        f.Controls.Add(btnSave);
        this.Controls.Add(f);
    }
}
