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

        public MainForm() { SetupUI(); }

        private void SetupUI()
        {
            this.Text = "G-Task Nexus | 整合通知中心";
            this.Size = new Size(650, 480);
            this.StartPosition = FormStartPosition.CenterScreen;

            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem configMenu = new ToolStripMenuItem("啟動選單 (System)");
            ToolStripMenuItem itemApi = new ToolStripMenuItem("API 憑證管理");
            itemApi.Click += (s, e) => OpenApiConfig();
            ToolStripMenuItem itemSync = new ToolStripMenuItem("重新整理同步");
            itemSync.Click += async (s, e) => await SyncTasks();
            ToolStripMenuItem itemAdd = new ToolStripMenuItem("新增項目");
            itemAdd.Click += async (s, e) => await PromptAddTask();

            configMenu.DropDownItems.AddRange(new ToolStripItem[] { itemApi, itemSync, itemAdd });
            menu.Items.Add(configMenu);

            _taskView = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true };
            _taskView.Columns.Add("狀態", 60);
            _taskView.Columns.Add("內容", 400);
            _taskView.Columns.Add("ID", 150);

            _status = new Label { Dock = DockStyle.Bottom, Height = 25, Text = " 就緒", BackColor = Color.WhiteSmoke };

            this.Controls.Add(_taskView);
            this.Controls.Add(_status);
            this.Controls.Add(menu);
            this.MainMenuStrip = menu;
        }

        private void OpenApiConfig()
        {
            using (Form f = new Form { Text = "貼入 API JSON", Size = new Size(400, 300) })
            {
                TextBox txt = new TextBox { Multiline = true, Dock = DockStyle.Fill, Text = CoreSecurity.LoadSecureData() };
                Button btn = new Button { Text = "儲存", Dock = DockStyle.Bottom, Height = 40 };
                btn.Click += (s, e) => { CoreSecurity.SaveSecureData(txt.Text); f.Close(); };
                f.Controls.Add(txt); f.Controls.Add(btn); f.ShowDialog();
            }
        }

        private async Task SyncTasks()
        {
            try {
                var items = await _service.GetAllTasksAsync();
                _taskView.Items.Clear();
                foreach (var t in items) {
                    var lvi = new ListViewItem(t.Status == "completed" ? "[V]" : "[ ]");
                    lvi.SubItems.Add(t.Title); lvi.SubItems.Add(t.Id);
                    _taskView.Items.Add(lvi);
                }
                _status.Text = " 同步完成";
            } catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private async Task PromptAddTask()
        {
            string title = Microsoft.VisualBasic.Interaction.InputBox("事項名稱:", "新增", "");
            if (!string.IsNullOrEmpty(title)) { await _service.AddTaskAsync(title); await SyncTasks(); }
        }
    }
}
