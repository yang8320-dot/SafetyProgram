using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace GTaskNexus
{
    public class MainForm : Form
    {
        private ListView _taskView;
        private Label _status;
        private GoogleTaskService _service = new GoogleTaskService();

        public MainForm()
        {
            SetupUI();
        }

        private void SetupUI()
        {
            this.Text = "G-Task Nexus | 整合通知中心";
            this.Size = new Size(650, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft JhengHei", 9);

            // 選單欄
            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem configMenu = new ToolStripMenuItem("啟動選單 (System)");
            
            ToolStripMenuItem itemApi = new ToolStripMenuItem("API 憑證管理 (新增/修改)");
            itemApi.Click += (s, e) => OpenApiConfig();
            
            ToolStripMenuItem itemSync = new ToolStripMenuItem("重新整理同步");
            itemSync.Click += async (s, e) => await SyncTasks();

            ToolStripMenuItem itemAdd = new ToolStripMenuItem("新增待辦項目");
            itemAdd.Click += async (s, e) => await PromptAddTask();

            configMenu.DropDownItems.AddRange(new ToolStripItem[] { itemApi, itemSync, itemAdd });
            menu.Items.Add(configMenu);

            // 任務清單
            _taskView = new ListView()
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            _taskView.Columns.Add("狀態", 60);
            _taskView.Columns.Add("待辦內容", 400);
            _taskView.Columns.Add("雲端索引", 150);

            // 狀態列
            _status = new Label() { Dock = DockStyle.Bottom, Height = 25, Text = " 系統就緒", BackColor = Color.WhiteSmoke };

            this.Controls.Add(_taskView);
            this.Controls.Add(_status);
            this.Controls.Add(menu);
            this.MainMenuStrip = menu;
        }

        private void OpenApiConfig()
        {
            using (Form f = new Form() { Text = "API Key 設定", Size = new Size(400, 300), StartPosition = FormStartPosition.CenterParent })
            {
                TextBox txt = new TextBox { Multiline = true, Dock = DockStyle.Fill, Text = CoreSecurity.LoadSecureData() };
                Button btn = new Button { Text = "加密並存檔", Dock = DockStyle.Bottom, Height = 40 };
                btn.Click += (s, e) => {
                    CoreSecurity.SaveSecureData(txt.Text);
                    MessageBox.Show("憑證已轉換為加密二進位檔！");
                    f.Close();
                };
                f.Controls.Add(txt);
                f.Controls.Add(btn);
                f.ShowDialog();
            }
        }

        private async Task SyncTasks()
        {
            UpdateStatus("正在同步 Google Tasks...");
            try {
                var items = await _service.GetAllTasksAsync();
                _taskView.Items.Clear();
                foreach (var t in items) {
                    var lvi = new ListViewItem(t.Status == "completed" ? "[V]" : "[ ]");
                    lvi.SubItems.Add(t.Title);
                    lvi.SubItems.Add(t.Id);
                    _taskView.Items.Add(lvi);
                }
                UpdateStatus("同步完成。");
            } catch (Exception ex) { MessageBox.Show(ex.Message); UpdateStatus("同步失敗"); }
        }

        private async Task PromptAddTask()
        {
            // 簡單輸入框 (需引用 Microsoft.VisualBasic)
            string title = Microsoft.VisualBasic.Interaction.InputBox("輸入新的待辦事項:", "G-Task Nexus", "");
            if (!string.IsNullOrEmpty(title)) {
                UpdateStatus("傳送至雲端...");
                await _service.AddTaskAsync(title);
                await SyncTasks();
            }
        }

        private void UpdateStatus(string msg)
        {
            if (this.InvokeRequired) this.Invoke(new Action(() => _status.Text = " " + msg));
            else _status.Text = " " + msg;
        }
    }
}
