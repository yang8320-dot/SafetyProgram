using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace GTaskNexus
{
    public class MainForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9527;
        private const int MOD_CONTROL = 0x0002;
        private const int WM_HOTKEY = 0x0312;

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
            this.Size = new Size(700, 550); // 稍微加寬視窗以容納日期欄位
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft JhengHei UI", 10.5F, FontStyle.Regular);
            this.BackColor = Color.White;

            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem configMenu = new ToolStripMenuItem("系統設定 (System)");
            ToolStripMenuItem itemApi = new ToolStripMenuItem("API 憑證管理");
            itemApi.Click += (s, e) => OpenApiConfig();
            configMenu.DropDownItems.Add(itemApi);
            menu.Items.Add(configMenu);

            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(245, 245, 245), Padding = new Padding(10) };
            
            _txtNewTask = new TextBox { Location = new Point(15, 15), Width = 480, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtNewTask.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await AddNewTask(); };

            _btnAdd = new Button { 
                Text = "+ 新增項目", Location = new Point(510, 14), Width = 120, Height = 30,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnAdd.FlatAppearance.BorderSize = 0;
            _btnAdd.Click += async (s, e) => await AddNewTask();

            pnlTop.Controls.Add(_txtNewTask); pnlTop.Controls.Add(_btnAdd);

            // --- 任務列表區塊 ---
            _taskView = new ListView { 
                Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, 
                GridLines = false, CheckBoxes = true, BorderStyle = BorderStyle.None,
                ShowGroups = true,
                LabelEdit = false // 關閉直接文字編輯，改為雙擊彈窗
            };
            
            // 新增欄位
            _taskView.Columns.Add("任務名稱 (雙擊進行進階編輯)", 450);
            _taskView.Columns.Add("到期日", 150);
            
            _groupTodo = new ListViewGroup("📌 待辦事項", HorizontalAlignment.Left);
            _groupDone = new ListViewGroup("✅ 已完成", HorizontalAlignment.Left);
            _taskView.Groups.Add(_groupTodo);
            _taskView.Groups.Add(_groupDone);

            // 事件綁定
            _taskView.ItemChecked += async (s, e) => await HandleTaskChecked(e);
            _taskView.DoubleClick += async (s, e) => await OpenEditDialog(); // 雙擊事件
            _taskView.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Delete) await DeleteSelectedTask(); };

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem deleteItem = new ToolStripMenuItem("🗑️ 刪除此任務");
            deleteItem.Click += async (s, e) => await DeleteSelectedTask();
            contextMenu.Items.Add(deleteItem);
            _taskView.ContextMenuStrip = contextMenu;

            _status = new Label { Dock = DockStyle.Bottom, Height = 25, Text = " 準備就緒", BackColor = Color.WhiteSmoke, ForeColor = Color.DimGray };

            this.Controls.Add(_taskView); this.Controls.Add(pnlTop); this.Controls.Add(_status); this.Controls.Add(menu);
            this.MainMenuStrip = menu;

            this.Shown += async (s, e) => await SyncTasksSilently();
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        // --- 動態建立「編輯視窗」 ---
        private async Task OpenEditDialog()
        {
            if (_taskView.SelectedItems.Count == 0) return;
            var item = _taskView.SelectedItems[0];
            string taskId = item.Tag?.ToString();
            if (string.IsNullOrEmpty(taskId)) return;

            // 判斷目前是否有日期
            DateTime? existingDate = null;
            if (item.SubItems.Count > 1 && DateTime.TryParse(item.SubItems[1].Text, out DateTime parsed))
            {
                existingDate = parsed;
            }

            // 建立彈出視窗
            using (Form editForm = new Form { Text = "編輯任務詳情", Size = new Size(380, 250), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                Label lblTitle = new Label { Text = "任務名稱:", Location = new Point(15, 15), AutoSize = true };
                TextBox txtTitle = new TextBox { Text = item.Text, Location = new Point(15, 40), Width = 330 };

                Label lblDate = new Label { Text = "設定到期日 (Google 僅支援日期，不含時間):", Location = new Point(15, 80), AutoSize = true, ForeColor = Color.Gray };
                
                CheckBox chkHasDate = new CheckBox { Text = "啟用日期", Location = new Point(15, 105), AutoSize = true, Checked = existingDate.HasValue };
                DateTimePicker dtpDate = new DateTimePicker { Location = new Point(100, 102), Width = 245, Format = DateTimePickerFormat.Short, Enabled = chkHasDate.Checked };
                
                if (existingDate.HasValue) dtpDate.Value = existingDate.Value;
                chkHasDate.CheckedChanged += (s, e) => dtpDate.Enabled = chkHasDate.Checked;

                Button btnSave = new Button { Text = "儲存變更", Location = new Point(130, 160), Width = 120, Height = 35, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                
                btnSave.Click += async (s, e) => {
                    if (string.IsNullOrWhiteSpace(txtTitle.Text)) return;
                    _autoSyncTimer.Stop();
                    _isSyncing = true;
                    btnSave.Enabled = false;
                    btnSave.Text = "儲存中...";

                    try {
                        DateTime? newDate = chkHasDate.Checked ? dtpDate.Value : (DateTime?)null;
                        await _service.UpdateTaskDetailsAsync(taskId, txtTitle.Text, newDate);
                        editForm.DialogResult = DialogResult.OK;
                    } catch {
                        MessageBox.Show("儲存失敗，請檢查網路連線。");
                        btnSave.Enabled = true;
                        btnSave.Text = "儲存變更";
                    } finally {
                        _isSyncing = false;
                        _autoSyncTimer.Start();
                    }
                };

                editForm.Controls.AddRange(new Control[] { lblTitle, txtTitle, lblDate, chkHasDate, dtpDate, btnSave });

                // 如果按下了儲存並成功關閉視窗，重新同步畫面
                if (editForm.ShowDialog() == DialogResult.OK)
                {
                    await SyncTasksSilently();
                }
            }
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

        private void SetupAutoSync()
        {
            _autoSyncTimer = new Timer { Interval = 15000 };
            _autoSyncTimer.Tick += async (s, e) => await SyncTasksSilently();
            _autoSyncTimer.Start();
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
                    _taskView.Items.Remove(item);
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
                    
                    // 解析雲端回傳的到期日
                    string dateStr = "";
                    if (!string.IsNullOrEmpty(t.Due))
                    {
                        if (DateTime.TryParse(t.Due, out DateTime parsedDate))
                        {
                            dateStr = parsedDate.ToString("yyyy/MM/dd");
                        }
                    }
                    lvi.SubItems.Add(dateStr); // 加入第二個欄位 (到期日)

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
