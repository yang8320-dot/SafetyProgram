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
        private CheckBox _chkNewTaskHasDate;
        private DateTimePicker _dtpNewTaskDate;
        private Button _btnAdd;
        private Button _btnClearDone;
        private CheckBox _chkShowCompleted;
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
            // 視窗加寬至 980，確保所有按鈕都有寬裕的空間
            this.Size = new Size(980, 600); 
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
            
            _txtNewTask = new TextBox { Location = new Point(15, 15), Width = 300, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtNewTask.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await AddNewTask(); };

            // 【修正重點 1】關閉 AutoSize，強制給定 65px 的寬度，絕不會再擋到後面的日期
            _chkNewTaskHasDate = new CheckBox { Text = "期限:", Location = new Point(330, 18), AutoSize = false, Width = 65, Cursor = Cursors.Hand };
            
            // 將日期選擇器向右移至安全位置 (X:395)
            _dtpNewTaskDate = new DateTimePicker { Location = new Point(395, 15), Width = 140, Format = DateTimePickerFormat.Short, Enabled = false };
            _chkNewTaskHasDate.CheckedChanged += (s, e) => _dtpNewTaskDate.Enabled = _chkNewTaskHasDate.Checked;

            _btnAdd = new Button { 
                Text = "+ 新增", Location = new Point(550, 14), Width = 80, Height = 30,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnAdd.FlatAppearance.BorderSize = 0;
            _btnAdd.Click += async (s, e) => await AddNewTask();

            _btnClearDone = new Button {
                Text = "🧹 清除已完成", Location = new Point(645, 14), Width = 130, Height = 30,
                BackColor = Color.FromArgb(204, 51, 51), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnClearDone.FlatAppearance.BorderSize = 0;
            _btnClearDone.Click += async (s, e) => await ClearAllCompletedTasks();

            _chkShowCompleted = new CheckBox {
                Text = "顯示已完成", Location = new Point(790, 18), AutoSize = true, Checked = false, Cursor = Cursors.Hand
            };
            _chkShowCompleted.CheckedChanged += async (s, e) => await SyncTasksSilently();

            pnlTop.Controls.AddRange(new Control[] { _txtNewTask, _chkNewTaskHasDate, _dtpNewTaskDate, _btnAdd, _btnClearDone, _chkShowCompleted });

            _taskView = new ListView { 
                Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, 
                GridLines = false, CheckBoxes = true, BorderStyle = BorderStyle.None,
                ShowGroups = true, LabelEdit = false
            };
            
            _taskView.Columns.Add("刪除", 60, HorizontalAlignment.Center);
            _taskView.Columns.Add("任務名稱 (雙擊進行進階編輯)", 560);
            _taskView.Columns.Add("到期日", 160);
            
            _groupTodo = new ListViewGroup("📌 待辦事項", HorizontalAlignment.Left);
            _groupDone = new ListViewGroup("✅ 已完成", HorizontalAlignment.Left);
            _taskView.Groups.Add(_groupTodo);
            _taskView.Groups.Add(_groupDone);

            _taskView.ItemCheck += TaskView_ItemCheck; 
            _taskView.ItemChecked += async (s, e) => await HandleTaskChecked(e);
            
            _taskView.MouseClick += async (s, e) => {
                if (e.Button != MouseButtons.Left) return;
                var hit = _taskView.HitTest(e.Location);
                if (hit.Item != null && hit.Item.SubItems.IndexOf(hit.SubItem) == 0 && hit.Location != ListViewHitTestLocations.StateImage) {
                    await DeleteTaskAction(hit.Item);
                }
            };

            _taskView.DoubleClick += async (s, e) => await OpenEditDialog();
            
            _taskView.KeyDown += async (s, e) => { 
                if (e.KeyCode == Keys.Delete && _taskView.SelectedItems.Count > 0) 
                    await DeleteTaskAction(_taskView.SelectedItems[0]); 
            };

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem deleteItem = new ToolStripMenuItem("🗑️ 刪除此任務");
            deleteItem.Click += async (s, e) => {
                if (_taskView.SelectedItems.Count > 0) await DeleteTaskAction(_taskView.SelectedItems[0]);
            };
            contextMenu.Items.Add(deleteItem);
            _taskView.ContextMenuStrip = contextMenu;

            _status = new Label { Dock = DockStyle.Bottom, Height = 25, Text = " 準備就緒", BackColor = Color.WhiteSmoke, ForeColor = Color.DimGray };

            this.Controls.Add(_taskView); this.Controls.Add(pnlTop); this.Controls.Add(_status); this.Controls.Add(menu);
            this.MainMenuStrip = menu;

            this.Shown += async (s, e) => await SyncTasksSilently();
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        private void TaskView_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_isSyncing) return;
            Point loc = _taskView.PointToClient(Cursor.Position);
            ListViewHitTestInfo hitTest = _taskView.HitTest(loc);
            if (hitTest.Location != ListViewHitTestLocations.StateImage) e.NewValue = e.CurrentValue; 
        }

        private async Task DeleteTaskAction(ListViewItem item)
        {
            if (item == null) return;
            string taskId = item.Tag?.ToString();
            string taskTitle = item.SubItems.Count > 1 ? item.SubItems[1].Text : "此任務";

            if (MessageBox.Show($"確定要快速刪除「{taskTitle}」嗎？", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                UpdateStatus(" 正在刪除任務...");
                try {
                    await _service.DeleteTaskAsync(taskId);
                    _taskView.Items.Remove(item);
                    UpdateStatus(" 任務已刪除。");
                } catch { UpdateStatus(" 刪除失敗。"); }
            }
        }

        private async Task ClearAllCompletedTasks()
        {
            if (MessageBox.Show("確定要永久刪除所有「已完成」的任務嗎？\n(此動作同步雲端且無法復原)", "一鍵清除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                UpdateStatus(" 正在清除雲端已完成任務...");
                try {
                    await _service.ClearCompletedTasksAsync();
                    await SyncTasksSilently();
                    UpdateStatus(" 清除完成！");
                } catch { UpdateStatus(" 清除失敗，請檢查網路連線。"); }
            }
        }

        private async Task OpenEditDialog()
        {
            var hit = _taskView.HitTest(_taskView.PointToClient(Cursor.Position));
            if (hit.Item != null && hit.Item.SubItems.IndexOf(hit.SubItem) == 0) return; 

            if (_taskView.SelectedItems.Count == 0) return;
            var item = _taskView.SelectedItems[0];
            string taskId = item.Tag?.ToString();
            if (string.IsNullOrEmpty(taskId)) return;

            DateTime? existingDate = null;
            if (item.SubItems.Count > 2 && DateTime.TryParse(item.SubItems[2].Text, out DateTime parsed)) existingDate = parsed;

            using (Form editForm = new Form { Text = "編輯任務詳情", Size = new Size(380, 250), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false })
            {
                Label lblTitle = new Label { Text = "任務名稱:", Location = new Point(15, 15), AutoSize = true };
                TextBox txtTitle = new TextBox { Text = item.SubItems[1].Text, Location = new Point(15, 40), Width = 330 };
                
                Label lblDate = new Label { Text = "設定到期日:", Location = new Point(15, 80), AutoSize = true, ForeColor = Color.Gray };
                CheckBox chkHasDate = new CheckBox { Text = "啟用日期", Location = new Point(15, 105), AutoSize = true, Checked = existingDate.HasValue };
                DateTimePicker dtpDate = new DateTimePicker { Location = new Point(100, 102), Width = 245, Format = DateTimePickerFormat.Short, Enabled = chkHasDate.Checked };
                
                if (existingDate.HasValue) dtpDate.Value = existingDate.Value;
                chkHasDate.CheckedChanged += (s, e) => dtpDate.Enabled = chkHasDate.Checked;

                Button btnSave = new Button { Text = "儲存變更", Location = new Point(130, 160), Width = 120, Height = 35, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                
                btnSave.Click += async (s, e) => {
                    if (string.IsNullOrWhiteSpace(txtTitle.Text)) return;
                    _autoSyncTimer.Stop(); _isSyncing = true; btnSave.Enabled = false; btnSave.Text = "儲存中...";
                    try {
                        DateTime? newDate = chkHasDate.Checked ? dtpDate.Value : (DateTime?)null;
                        await _service.UpdateTaskDetailsAsync(taskId, txtTitle.Text, newDate);
                        editForm.DialogResult = DialogResult.OK;
                    } catch {
                        MessageBox.Show("儲存失敗，請檢查網路。"); btnSave.Enabled = true; btnSave.Text = "儲存變更";
                    } finally { _isSyncing = false; _autoSyncTimer.Start(); }
                };

                editForm.AcceptButton = btnSave;
                editForm.Controls.AddRange(new Control[] { lblTitle, txtTitle, lblDate, chkHasDate, dtpDate, btnSave });
                if (editForm.ShowDialog() == DialogResult.OK) await SyncTasksSilently();
            }
        }

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

        // 【修正重點 2】處理打勾狀態變更時的字體顏色
        private async Task HandleTaskChecked(ItemCheckedEventArgs e)
        {
            if (_isSyncing) return;
            string taskId = e.Item.Tag?.ToString();
            bool isDone = e.Item.Checked;

            if (!string.IsNullOrEmpty(taskId))
            {
                string taskTitle = e.Item.SubItems.Count > 1 ? e.Item.SubItems[1].Text : "";
                UpdateStatus($"正在將 '{taskTitle}' 狀態同步至雲端...");
                
                // 確保第 0 欄 (❌) 永遠是紅色的
                e.Item.UseItemStyleForSubItems = false; 
                e.Item.ForeColor = Color.Red; 
                
                // 根據打勾狀態設定後面文字的顏色與刪除線
                Color textColor = isDone ? Color.Gray : Color.Black;
                Font textFont = new Font(_taskView.Font, isDone ? FontStyle.Strikeout : FontStyle.Regular);

                if (e.Item.SubItems.Count > 1) {
                    e.Item.SubItems[1].ForeColor = textColor;
                    e.Item.SubItems[1].Font = textFont;
                }
                if (e.Item.SubItems.Count > 2) {
                    e.Item.SubItems[2].ForeColor = textColor;
                    e.Item.SubItems[2].Font = textFont;
                }

                e.Item.Group = isDone ? _groupDone : _groupTodo;

                try {
                    await _service.UpdateTaskStatusAsync(taskId, isDone);
                    UpdateStatus(" 狀態同步完成。");
                } catch { UpdateStatus(" 狀態同步失敗，請檢查網路。"); }
            }
        }

        private async Task AddNewTask()
        {
            string title = _txtNewTask.Text.Trim();
            if (string.IsNullOrEmpty(title)) return;
            _txtNewTask.Enabled = false; _btnAdd.Enabled = false;
            UpdateStatus(" 正在新增至 Google Tasks...");

            try {
                DateTime? dueDate = _chkNewTaskHasDate.Checked ? _dtpNewTaskDate.Value : (DateTime?)null;
                await _service.AddTaskAsync(title, dueDate);
                
                _txtNewTask.Clear();
                _chkNewTaskHasDate.Checked = false;
                await SyncTasksSilently();
            } catch (Exception ex) { MessageBox.Show("新增失敗: " + ex.Message); } 
            finally { _txtNewTask.Enabled = true; _btnAdd.Enabled = true; _txtNewTask.Focus(); }
        }

        private async Task SyncTasksSilently()
        {
            if (_isSyncing) return;
            _isSyncing = true;
            UpdateStatus(" 正在與雲端同步...");

            try {
                var items = await _service.GetAllTasksAsync(_chkShowCompleted.Checked);
                
                _taskView.BeginUpdate();
                _taskView.Items.Clear();
                
                foreach (var t in items) {
                    var lvi = new ListViewItem(" ❌"); 
                    lvi.Tag = t.Id; 
                    lvi.Checked = (t.Status == "completed");

                    // 【修正重點 3】關閉子項目的統一風格，強制把第 0 欄變成紅色
                    lvi.UseItemStyleForSubItems = false;
                    lvi.ForeColor = Color.Red; 
                    
                    var subTitle = lvi.SubItems.Add(t.Title ?? "無標題"); 

                    string dateStr = "";
                    if (!string.IsNullOrEmpty(t.Due) && DateTime.TryParse(t.Due, out DateTime parsedDate)) {
                        dateStr = parsedDate.ToString("yyyy/MM/dd");
                    }
                    var subDate = lvi.SubItems.Add(dateStr); 

                    // 針對後面的欄位動態套用顏色與字體
                    Color rowColor = lvi.Checked ? Color.Gray : Color.Black;
                    Font rowFont = new Font(_taskView.Font, lvi.Checked ? FontStyle.Strikeout : FontStyle.Regular);
                    
                    subTitle.ForeColor = rowColor;
                    subTitle.Font = rowFont;
                    subDate.ForeColor = rowColor;
                    subDate.Font = rowFont;

                    lvi.Group = lvi.Checked ? _groupDone : _groupTodo;
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
