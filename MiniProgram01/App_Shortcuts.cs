/*
 * 檔案功能：常用捷徑管理與開啟模組 (Microsoft.Data.Sqlite 升級版)
 */

using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

public class App_Shortcuts : UserControl
{
    private MainForm parentForm;
    private FlowLayoutPanel taskPanel;
    private int dragInsertIndex = -1;

    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleGreen = Color.FromArgb(52, 199, 89);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    public class ShortcutItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
    }

    public List<ShortcutItem> shortcuts = new List<ShortcutItem>();

    public App_Shortcuts(MainForm mainForm)
    {
        this.parentForm = mainForm;
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15);

        InitializeUI();
        _ = LoadShortcutsAsync();
    }

    private void InitializeUI()
    {
        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = 45, ColumnCount = 2, BackColor = Color.Transparent };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));

        Label lblTitle = new Label() { Text = "常用捷徑", Font = new Font(MainFont.FontFamily, 14f, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5, 0, 0, 0) };
        Button btnAdd = new Button() { Text = "新增", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = AppleBlue, ForeColor = Color.White, Font = BoldFont, Margin = new Padding(0, 5, 0, 5) };
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += (s, e) => { new EditShortcutWindow(this, -1, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);

        taskPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = AppleBgColor, AllowDrop = true };
        taskPanel.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskPanel.DragOver += OnTaskDragOver;
        taskPanel.DragLeave += (s, e) => { dragInsertIndex = -1; taskPanel.Invalidate(); };
        taskPanel.DragDrop += OnTaskDragDrop;
        taskPanel.Paint += OnTaskContainerPaint;

        taskPanel.Resize += (s, e) =>
        {
            int safeWidth = taskPanel.ClientSize.Width - 20;
            if (safeWidth > 0) { foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth; }
        };

        this.Controls.Add(taskPanel);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent });
        this.Controls.Add(header);
        taskPanel.BringToFront();
    }

    public void RefreshUI()
    {
        if (this.InvokeRequired) { this.Invoke(new Action(RefreshUI)); return; }

        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > 50 ? taskPanel.ClientSize.Width - 20 : 450;

        foreach (var s in shortcuts)
        {
            Panel card = new Panel() { Width = startWidth, AutoSize = true, MinimumSize = new Size(0, 55), Margin = new Padding(0, 0, 0, 15), BackColor = Color.White, Padding = new Padding(10) };
            TableLayoutPanel tlp = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, AutoSize = true };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f)); 

            Button btnOpen = new Button() { Text = "啟動", Dock = DockStyle.Top, Height = 35, BackColor = AppleGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = BoldFont, Margin = new Padding(0, 0, 10, 0) };
            btnOpen.FlatAppearance.BorderSize = 0;
            btnOpen.Click += (sender, e) => { try { Process.Start(new ProcessStartInfo() { FileName = s.Path, UseShellExecute = true }); } catch { MessageBox.Show("無法開啟此捷徑，請檢查路徑是否存在！", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error); } };

            Button btnDel = new Button() { Text = "刪除", Dock = DockStyle.Top, Height = 35, BackColor = AppleRed, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = BoldFont, Margin = new Padding(0, 0, 10, 0) };
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

            Label lbl = new Label() { Text = s.Name, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Font = MainFont, Padding = new Padding(10, 0, 10, 0), Cursor = Cursors.SizeAll };
            lbl.MouseDown += (sender, e) => { if (e.Button == MouseButtons.Left) card.DoDragDrop(card, DragDropEffects.Move); };

            Button btnEdit = new Button() { Text = "編輯", Dock = DockStyle.Top, Height = 35, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = BoldFont, Margin = new Padding(10, 0, 0, 0) };
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (sender, e) => { new EditShortcutWindow(this, shortcuts.IndexOf(s), s).ShowDialog(); };

            tlp.Controls.Add(btnOpen, 0, 0); tlp.Controls.Add(btnDel, 1, 0); tlp.Controls.Add(lbl, 2, 0); tlp.Controls.Add(btnEdit, 3, 0);
            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }

    private void OnTaskDragOver(object sender, DragEventArgs e)
    {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskPanel.PointToClient(new Point(e.X, e.Y));
        Control target = taskPanel.GetChildAtPoint(clientPoint);
        
        if (target != null && target is Panel)
        {
            int idx = taskPanel.Controls.GetChildIndex(target);
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            if (dragInsertIndex != idx) { dragInsertIndex = idx; taskPanel.Invalidate(); }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e)
    {
        if (dragInsertIndex != -1 && taskPanel.Controls.Count > 0)
        {
            int y = (dragInsertIndex < taskPanel.Controls.Count) ? taskPanel.Controls[dragInsertIndex].Top - 7 : taskPanel.Controls[taskPanel.Controls.Count - 1].Bottom + 7;
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
            
            var item = shortcuts[currentIdx];
            shortcuts.RemoveAt(currentIdx);
            int finalIdx = Math.Min(targetIdx, shortcuts.Count);
            shortcuts.Insert(finalIdx, item);

            dragInsertIndex = -1; 
            taskPanel.Invalidate(); 
            
            await SaveShortcutsAsync(); 
            RefreshUI();
        }
    }

    public async Task LoadShortcutsAsync()
    {
        try
        {
            var list = await Task.Run(() =>
            {
                var tempList = new List<ShortcutItem>();
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Id, Name, TargetPath FROM Shortcuts ORDER BY SortOrder ASC", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tempList.Add(new ShortcutItem { Id = reader.GetString(0), Name = reader.GetString(1), Path = reader.GetString(2) });
                        }
                    }
                }
                return tempList;
            });

            shortcuts = list;
            RefreshUI();
        }
        catch (Exception ex) { MessageBox.Show($"載入捷徑失敗: {ex.Message}"); }
    }

    public async Task SaveShortcutsAsync()
    {
        try
        {
            await Task.Run(() =>
            {
                using (var conn = DatabaseManager.GetConnection())
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        using (var cmdDel = new SqliteCommand("DELETE FROM Shortcuts", conn, transaction)) { cmdDel.ExecuteNonQuery(); }

                        using (var cmdIns = new SqliteCommand("INSERT INTO Shortcuts (Id, Name, TargetPath, SortOrder) VALUES (@Id, @Name, @Path, @Order)", conn, transaction))
                        {
                            for (int i = 0; i < shortcuts.Count; i++)
                            {
                                cmdIns.Parameters.Clear();
                                cmdIns.Parameters.AddWithValue("@Id", string.IsNullOrEmpty(shortcuts[i].Id) ? Guid.NewGuid().ToString("N") : shortcuts[i].Id);
                                cmdIns.Parameters.AddWithValue("@Name", shortcuts[i].Name);
                                cmdIns.Parameters.AddWithValue("@Path", shortcuts[i].Path);
                                cmdIns.Parameters.AddWithValue("@Order", i);
                                cmdIns.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                    }
                }
            });
        }
        catch (Exception ex) { MessageBox.Show($"儲存捷徑失敗: {ex.Message}"); }
    }
}

public class EditShortcutWindow : Form
{
    private App_Shortcuts parent;
    private int index;
    private TextBox txtName, txtPath;

    public EditShortcutWindow(App_Shortcuts p, int idx, App_Shortcuts.ShortcutItem item)
    {
        this.parent = p; this.index = idx;
        
        this.Text = idx == -1 ? "新增捷徑" : "編輯捷徑";
        this.Width = 450; this.Height = 300;
        this.StartPosition = FormStartPosition.CenterScreen; this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false; this.MinimizeBox = false; this.BackColor = Color.FromArgb(245, 245, 247);
        this.AutoScaleMode = AutoScaleMode.Dpi; this.Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Regular);

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20) };

        f.Controls.Add(new Label() { Text = "捷徑名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        txtName = new TextBox() { Width = 380, Text = item?.Name ?? "", Margin = new Padding(0, 0, 0, 15), BorderStyle = BorderStyle.FixedSingle };
        f.Controls.Add(txtName);

        f.Controls.Add(new Label() { Text = "目標路徑 (檔案 / 資料夾 / 網址)：", AutoSize = true, Margin = new Padding(0, 0, 0, 5) });
        TableLayoutPanel pathRow = new TableLayoutPanel() { Width = 380, Height = 35, ColumnCount = 2, Margin = new Padding(0, 0, 0, 20) };
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));
        
        txtPath = new TextBox() { Dock = DockStyle.Fill, Text = item?.Path ?? "", BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 5, 10, 0) };
        Button btnBrowse = new Button() { Text = "瀏覽", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = Color.White };
        btnBrowse.FlatAppearance.BorderColor = Color.Gray;
        btnBrowse.Click += (s, e) => { OpenFileDialog ofd = new OpenFileDialog() { Title = "選擇捷徑目標" }; if (ofd.ShowDialog() == DialogResult.OK) { txtPath.Text = ofd.FileName; } };

        pathRow.Controls.Add(txtPath, 0, 0); pathRow.Controls.Add(btnBrowse, 1, 0); f.Controls.Add(pathRow);

        Button btnSave = new Button() { Text = "儲存設定", Width = 380, Height = 40, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Microsoft JhengHei UI", 10f, FontStyle.Bold) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += async (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtPath.Text)) { MessageBox.Show("名稱與路徑不可為空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (index == -1) parent.shortcuts.Add(new App_Shortcuts.ShortcutItem() { Id = Guid.NewGuid().ToString("N"), Name = txtName.Text, Path = txtPath.Text });
            else { parent.shortcuts[index].Name = txtName.Text; parent.shortcuts[index].Path = txtPath.Text; }

            await parent.SaveShortcutsAsync();
            parent.RefreshUI();
            this.Close();
        };
        
        f.Controls.Add(btnSave); this.Controls.Add(f);
    }
}
