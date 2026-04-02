using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace GTaskNexus
{
    public class MainForm : Form
    {
        // --- 快捷鍵底層 API ---
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9527;
        private const int MOD_CONTROL = 0x0002;
        private const int WM_HOTKEY = 0x0312;

        // --- 控制項宣告 ---
        private ListView _taskView;
        private Label _status;
        private TextBox _txtNewTask;
        private Button _btnAdd;
        private Timer _autoSyncTimer;
        
        private ListViewGroup _groupTodo;
        private ListViewGroup _groupDone;

        private GoogleTaskService _service = new GoogleTaskService();
        private bool _isSyncing = false;

        public MainForm() 
        { 
            SetupUI();
            SetupAutoSync();
        }

        private void SetupUI()
        {
            this.Text = "G-Task Nexus | 雙向同步中心 (快捷鍵: Ctrl+2)";
            this.Size = new Size(650, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft JhengHei UI", 10.5F, FontStyle.Regular);
            this.BackColor = Color.White;

            // 頂部選單
            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem configMenu = new ToolStripMenuItem("系統設定 (System)");
            ToolStripMenuItem itemApi = new ToolStripMenuItem("API 憑證管理");
            itemApi.Click += (s, e) => OpenApiConfig();
            configMenu.DropDownItems.Add(itemApi);
            menu.Items.Add(configMenu);

            // 上方新增區塊
            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(245, 245, 245), Padding = new Padding(10) };
            
            _txtNewTask = new TextBox { Location = new Point(15, 15), Width = 450, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtNewTask.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await AddNewTask(); };

            _btnAdd = new Button { 
                Text = "+ 新增項目", Location = new Point(480, 14), Width = 120, Height = 30,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnAdd.FlatAppearance.BorderSize = 0;
            _btnAdd.Click += async (s, e) => await AddNewTask();

            pnlTop.Controls.Add(_txtNewTask); pnlTop.Controls.Add(_btnAdd);

            // --- 任務列表區塊 (加入群組與編輯功能) ---
            _taskView = new ListView { 
                Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, 
                GridLines = false, CheckBoxes = true, BorderStyle = BorderStyle.None,
                LabelEdit = true, // 允許直接點擊文字修改
                ShowGroups = true // 啟用群組分段顯示
            };
            _taskView.Columns.Add("任務名稱 (點擊兩下可編修)", 550);
            
            // 建立兩個群組
            _groupTodo = new ListViewGroup("📌 待辦事項", HorizontalAlignment.Left);
            _groupDone = new ListViewGroup("✅ 已完成", HorizontalAlignment.Left);
            _taskView.Groups.Add(_groupTodo);
            _taskView.Groups.Add(_groupDone);

            // 列表事件綁定
            _taskView.ItemChecked += async (s, e) => await HandleTaskChecked(e);
            _taskView.BeforeLabelEdit += TaskView_BeforeLabelEdit;
            _taskView.AfterLabelEdit += TaskView_AfterLabelEdit;
            _taskView.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Delete) await DeleteSelectedTask(); };

            // 右鍵刪除選單
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem deleteItem = new ToolStripMenuItem("🗑️ 刪除此任務");
            deleteItem.Click += async (s, e) => await DeleteSelectedTask();
            contextMenu.Items.Add(deleteItem);
            _taskView.ContextMenuStrip = contextMenu;

            // 底部狀態列
            _status = new Label { Dock = DockStyle.Bottom, Height = 25, Text = " 準備就緒", BackColor = Color.WhiteSmoke, ForeColor = Color.DimGray };

            this.Controls.Add(_taskView); this.Controls.Add(pnlTop); this.Controls.Add(_status); this.Controls.Add(menu);
            this.MainMenuStrip = menu;

            this.Shown += async (s, e) => await SyncTasksSilently();
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        // --- 快捷鍵邏輯 ---
        private void MainForm_Load(object sender, EventArgs e) { RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, (int)Keys.D2); }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) { UnregisterHotKey(this.Handle, HOTKEY_ID); }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID) {
                if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
                this.Show(); this.Activate(); _txtNewTask.Focus();
            }
            base.WndProc(ref m);
        }

        // --- 同步與功能邏輯 ---
        private void SetupAutoSync()
        {
            _autoSyncTimer = new Timer { Interval = 15000 };
            _autoSyncTimer.Tick += async (s, e) => await SyncTasksSilently();
            _autoSyncTimer.Start();
        }

        // 開始編輯前：暫停計時器，以免打字被打斷
        private void TaskView_BeforeLabelEdit(object sender, LabelEditEventArgs e)
        {
            _autoSyncTimer.Stop();
            _isSyncing = true; // 鎖定狀態
        }

        // 結束編輯後：將新文字推送到雲端
        private async void TaskView_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
            _isSyncing = false;
            _autoSyncTimer.Start();

            // e.Label 為 null 代表使用者取消編輯 (例如按下 Esc)
            if (e.Label != null && !string.IsNullOrWhiteSpace(e.Label))
            {
                string taskId = _taskView.Items[e.Item].Tag?.ToString();
                if (taskId != null)
                {
                    UpdateStatus(" 正在更新任務名稱...");
                    try {
                        await _service.UpdateTaskTitleAsync(taskId, e.Label);
                        UpdateStatus(" 名稱更新完成。");
                    } catch {
                        e.CancelEdit = true; // 發生錯誤時還原文字
                        UpdateStatus(" 更新失敗，請檢查網路。");
                    }
                }
            }
            else if (e.Label != null && string.IsNullOrWhiteSpace(e.Label))
            {
                e.CancelEdit = true; // 不允許空白標題
            }
        }

        private async Task DeleteSelectedTask()
        {
            if (_taskView.SelectedItems.Count == 0) return;
            
            var item = _taskView.SelectedItems[0];
            string taskId = item.Tag?.ToString();

            if (MessageBox.Show($"確定要刪除「{item.Text}」嗎？", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                UpdateStatus(" 正在刪除任務...");
                try {
                    await _service.DeleteTaskAsync(taskId);
                    _taskView.Items.Remove(item); // 直接從畫面上移除，不需要等下次同步
                    UpdateStatus(" 任務已刪除。");
                } catch {
                    UpdateStatus(" 刪除失敗。");
                }
            }
        }

        private async Task HandleTaskChecked(ItemCheckedEventArgs e)
        {
            if (_isSyncing) return;
            string taskId = e.Item.Tag?.ToString();
            bool isDone = e.Item.Checked;

            if (!string.IsNullOrEmpty(taskId))
            {
                UpdateStatus($"正在將 '{e.Item.Text}' 狀態同步至雲端...");
                e.Item.ForeColor = isDone ? Color.Gray : Color.Black;
                e.Item.Font = new Font(e.Item.Font, isDone ? FontStyle.Strikeout : FontStyle.Regular);
                
                // 動態移動群組
                e.Item.Group = isDone ? _groupDone : _groupTodo;

                try {
                    await _service.UpdateTaskStatusAsync(taskId, isDone);
                    UpdateStatus(" 狀態同步完成。");
                } catch {
                    UpdateStatus(" 狀態同步失敗，請檢查網路。");
                }
            }
        }

        private async Task AddNewTask()
        {
            string title = _txtNewTask.Text.Trim();
            if (string.IsNullOrEmpty(title)) return;

            _txtNewTask.Enabled = false; _btnAdd.Enabled = false;
            UpdateStatus(" 正在新增至 Google Tasks...");

            try {
                await _service.AddTaskAsync(title);
                _txtNewTask.Clear();
                await SyncTasksSilently();
            } catch (Exception ex) {
                MessageBox.Show("新增失敗: " + ex.Message);
            } finally {
                _txtNewTask.Enabled = true; _btnAdd.Enabled = true; _txtNewTask.Focus();
            }
        }

        private async Task SyncTasksSilently()
        {
            if (_isSyncing) return;
            _isSyncing = true;
            UpdateStatus(" 正在與雲端同步...");

            try {
                var items = await _service.GetAllTasksAsync();
                
                _taskView.BeginUpdate();
                _taskView.Items.Clear();
                
                foreach (var t in items) {
                    var lvi = new ListViewItem(t.Title ?? "無標題");
                    lvi.Tag = t.Id; 
                    lvi.Checked = (t.Status == "completed");
                    
                    // 根據狀態分配到不同的群組
                    lvi.Group = lvi.Checked ? _groupDone : _groupTodo;
                    
                    if (lvi.Checked) {
                        lvi.ForeColor = Color.Gray;
                        lvi.Font = new Font(_taskView.Font, FontStyle.Strikeout);
                    }
                    _taskView.Items.Add(lvi);
                }
                _taskView.EndUpdate();
                UpdateStatus($" 同步完成。最後更新: {DateTime.Now:HH:mm:ss}");
            } 
            catch { UpdateStatus(" 背景同步失敗，請確認 API 或網路連線。"); }
            finally { _isSyncing = false; }
        }

        private void OpenApiConfig()
        {
            // ... (維持原樣) ...
            using (Form f = new Form { Text = "貼入 API JSON", Size = new Size(400, 300), StartPosition = FormStartPosition.CenterParent })
            {
                TextBox txt = new TextBox { Multiline = true, Dock = DockStyle.Fill, Text = CoreSecurity.LoadSecureData() };
                Button btn = new Button { Text = "儲存並重新啟動", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.FromArgb(45,45,45), ForeColor=Color.White };
                btn.Click += (s, e) => { CoreSecurity.SaveSecureData(txt.Text); Application.Restart(); };
                f.Controls.Add(txt); f.Controls.Add(btn); f.ShowDialog();
            }
        }

        private void UpdateStatus(string msg)
        {
            if (this.InvokeRequired) this.Invoke(new Action(() => _status.Text = msg));
            else _status.Text = msg;
        }
    }
}
