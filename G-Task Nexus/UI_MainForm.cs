using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Collections.Generic;
// 注意：我們不在這裡直接引用 Google.Apis.Tasks.v1.Data 以避免與 System.Threading.Tasks.Task 衝突

namespace GTaskNexus
{
    public class MainForm : Form
    {
        // 定義控制項
        private ListView _taskView;
        private Label _statusLabel;
        private MenuStrip _mainMenu;
        
        // 引用邏輯服務層
        private GoogleTaskService _apiService = new GoogleTaskService();

        public MainForm()
        {
            // 初始化介面
            InitializeComponent();
            
            // 視窗載入後的小提示
            this.Load += (s, e) => UpdateStatus(" 系統就緒。請從選單進行 API 設定或同步。");
        }

        /// <summary>
        /// Code-First UI 佈局設定 (不使用 Designer)
        /// </summary>
        private void InitializeComponent()
        {
            // 1. 視窗主體設定
            this.Text = "G-Task Nexus | 雲端同步中心";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft JhengHei", 9);
            this.BackColor = Color.White;

            // 2. 建立選單 (Startup Menu)
            _mainMenu = new MenuStrip();
            ToolStripMenuItem menuSystem = new ToolStripMenuItem("系統控制 (&S)");
            
            ToolStripMenuItem itemApi = new ToolStripMenuItem("API 憑證管理 (新增/修改)");
            itemApi.ShortcutKeys = Keys.Control | Keys.P;
            itemApi.Click += (s, e) => ShowApiConfigDialog();

            ToolStripMenuItem itemSync = new ToolStripMenuItem("立即同步雲端資料");
            itemSync.ShortcutKeys = Keys.F5;
            itemSync.Click += async (s, e) => await SyncTasksAsync();

            ToolStripMenuItem itemAdd = new ToolStripMenuItem("新增待辦事項");
            itemAdd.ShortcutKeys = Keys.Control | Keys.N;
            itemAdd.Click += async (s, e) => await PromptAddTaskAsync();

            menuSystem.DropDownItems.AddRange(new ToolStripItem[] { itemApi, itemSync, itemAdd });
            _mainMenu.Items.Add(menuSystem);

            // 3. 建立任務列表 (ListView)
            _taskView = new ListView()
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("Microsoft JhengHei", 10),
                BorderStyle = BorderStyle.None
            };
            _taskView.Columns.Add("狀態", 70, HorizontalAlignment.Center);
            _taskView.Columns.Add("待辦內容描述", 420, HorizontalAlignment.Left);
            _taskView.Columns.Add("雲端唯一識別碼", 150, HorizontalAlignment.Left);

            // 4. 建立狀態列
            _statusLabel = new Label()
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.DimGray,
                Text = " 正在初始化..."
            };

            // 將控制項加入視窗
            this.Controls.Add(_taskView);
            this.Controls.Add(_statusLabel);
            this.Controls.Add(_mainMenu);
            this.MainMenuStrip = _mainMenu;
        }

        /// <summary>
        /// 彈出 API 設定視窗
        /// </summary>
        private void ShowApiConfigDialog()
        {
            using (Form configForm = new Form())
            {
                configForm.Text = "API 憑證管理 (貼入 JSON 內容)";
                configForm.Size = new Size(450, 350);
                configForm.StartPosition = FormStartPosition.CenterParent;
                configForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                configForm.MaximizeBox = false;

                TextBox txtJson = new TextBox { 
                    Multiline = true, 
                    Dock = DockStyle.Top, 
                    Height = 230, 
                    ScrollBars = ScrollBars.Vertical,
                    Text = CoreSecurity.LoadSecureData() // 自動載入已存的加密資料
                };

                Button btnSave = new Button { 
                    Text = "加密儲存並套用設定", 
                    Dock = DockStyle.Bottom, 
                    Height = 50,
                    BackColor = Color.FromArgb(0, 122, 204),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };

                btnSave.Click += (s, e) => {
                    if (string.IsNullOrWhiteSpace(txtJson.Text)) {
                        MessageBox.Show("內容不可為空！");
                        return;
                    }
                    CoreSecurity.SaveSecureData(txtJson.Text);
                    MessageBox.Show("憑證已成功轉換為隨機二進位加密檔！", "G-Task Nexus");
                    configForm.Close();
                };

                configForm.Controls.Add(txtJson);
                configForm.Controls.Add(btnSave);
                configForm.ShowDialog();
            }
        }

        /// <summary>
        /// 同步任務邏輯 (解決 Task 衝突)
        /// </summary>
        private async Task SyncTasksAsync()
        {
            UpdateStatus(" 正在連接 Google API 並下載任務...");
            _taskView.Items.Clear();

            try
            {
                // 調用服務層獲取資料 (這裡 Service 已經處理好 GData.Task 衝突)
                var tasks = await _apiService.GetAllTasksAsync();

                foreach (var t in tasks)
                {
                    string statusIcon = t.Status == "completed" ? "[V] 已完成" : "[ ] 待處理";
                    ListViewItem item = new ListViewItem(statusIcon);
                    item.SubItems.Add(t.Title ?? "(無標題)");
                    item.SubItems.Add(t.Id ?? "---");

                    if (t.Status == "completed") item.ForeColor = Color.Gray;
                    _taskView.Items.Add(item);
                }

                UpdateStatus($" 同步成功！共計 {tasks.Count} 個項目。同步時間: {DateTime.Now:HH:mm:ss}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("同步發生錯誤: " + ex.Message, "連線失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus(" 同步中斷，請檢查 API 設定與網路。");
            }
        }

        /// <summary>
        /// 新增任務彈窗
        /// </summary>
        private async Task PromptAddTaskAsync()
        {
            // 使用 Microsoft.VisualBasic 的 InputBox
            string taskTitle = Microsoft.VisualBasic.Interaction.InputBox(
                "請輸入欲新增至 Google Tasks 的事項名稱：", 
                "新增待辦項目", 
                "新任務");

            if (!string.IsNullOrWhiteSpace(taskTitle))
            {
                UpdateStatus(" 正在推送新項目至雲端...");
                try
                {
                    await _apiService.AddTaskAsync(taskTitle);
                    await SyncTasksAsync(); // 新增完畢自動重新整理
                }
                catch (Exception ex)
                {
                    MessageBox.Show("新增失敗: " + ex.Message);
                    UpdateStatus(" 新增作業失敗。");
                }
            }
        }

        /// <summary>
        /// 執行緒安全地更新狀態列
        /// </summary>
        private void UpdateStatus(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => _statusLabel.Text = message));
            }
            else
            {
                _statusLabel.Text = message;
            }
        }
    }
}
