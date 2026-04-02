using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace GTaskNexus
{
    public class MainForm : Form
    {
        // --- 呼叫 Windows 底層 API 註冊全域快捷鍵 ---
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID = 9527; // 自訂一個不衝突的快捷鍵 ID
        private const int MOD_CONTROL = 0x0002; // Ctrl 鍵的修飾碼
        private const int WM_HOTKEY = 0x0312; // Windows 快捷鍵訊息代碼

        // --- 控制項宣告 ---
        private ListView _taskView;
        private Label _status;
        private TextBox _txtNewTask;
        private Button _btnAdd;
        private Timer _autoSyncTimer;
        
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

            // 任務列表區塊 (含 CheckBox)
            _taskView = new ListView { 
                Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, 
                GridLines = false, CheckBoxes = true, BorderStyle = BorderStyle.None
            };
            _taskView.Columns.Add("任務名稱", 500);
            _taskView.ItemChecked += async (s, e) => await HandleTaskChecked(e);

            // 底部狀態列
            _status = new Label { Dock = DockStyle.Bottom, Height = 25, Text = " 準備就緒", BackColor = Color.WhiteSmoke, ForeColor = Color.DimGray };

            this.Controls.Add(_taskView); this.Controls.Add(pnlTop); this.Controls.Add(_status); this.Controls.Add(menu);
            this.MainMenuStrip = menu;

            // 視窗事件綁定
            this.Shown += async (s, e) => await SyncTasksSilently();
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        // --- 快捷鍵註冊與捕捉 ---
        private void MainForm_Load(object sender, EventArgs e)
        {
            // 註冊全域快捷鍵：Ctrl + 2 (Keys.D2 代碼為 0x32)
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, (int)Keys.D2);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 程式關閉時必須註銷快捷鍵，釋放系統資源
            UnregisterHotKey(this.Handle, HOTKEY_ID);
        }

        // 複寫視窗訊息處理機制，攔截我們設定的快捷鍵
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                WakeUpAndFocus();
            }
            base.WndProc(ref m);
        }

        // 喚醒視窗並對焦輸入框
        private void WakeUpAndFocus()
        {
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal;
            
            this.Show();
            this.Activate(); // 移至最上層
            _txtNewTask.Focus(); // 游標跳到新增欄位
        }
        // --- 快捷鍵邏輯結束 ---

        private void SetupAutoSync()
        {
            _autoSyncTimer = new Timer { Interval = 15000 };
            _autoSyncTimer.Tick += async (s, e) => await SyncTasksSilently();
            _autoSyncTimer.Start();
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
