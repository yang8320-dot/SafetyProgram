using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Drawing.Printing;

namespace GTaskNexus
{
    public class MainForm : Form
    {
        private ListView _taskView;
        private Label _status;
        private TextBox _txtNewTask;
        private CheckBox _chkNewTaskHasDate;
        private DateTimePicker _dtpNewTaskDate;
        private Button _btnAdd;
        private Button _btnClearDone;
        private Button _btnExportPdf;
        private CheckBox _chkShowCompleted;
        private Timer _autoSyncTimer;
        
        private ListViewGroup _groupTodo;
        private ListViewGroup _groupDone;

        private GoogleTaskService _service = new GoogleTaskService();
        private bool _isSyncing = false;
        
        private int _printGroupIndex = 0;
        private int _printItemIndex = 0;

        public MainForm() 
        { 
            SetupUI();
            SetupAutoSync();
        }

        private void SetupUI()
        {
            // 注意：此標題必須與 Program.cs 中的 FindWindow 字串完全相同
            this.Text = "G-Task Nexus | 雙向同步中心";
            this.Size = new Size(880, 650); 
            this.MinimumSize = new Size(800, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft JhengHei UI", 10.5F, FontStyle.Regular);
            this.BackColor = Color.White;

            MenuStrip menu = new MenuStrip();
            ToolStripMenuItem configMenu = new ToolStripMenuItem("系統設定 (System)");
            ToolStripMenuItem itemApi = new ToolStripMenuItem("API 憑證管理");
            itemApi.Click += (s, e) => OpenApiConfig();
            configMenu.DropDownItems.Add(itemApi);
            menu.Items.Add(configMenu);

            Panel pnlTopContainer = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = Color.FromArgb(245, 245, 245) };

            FlowLayoutPanel pnlRow1 = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 45, Padding = new Padding(10, 10, 10, 0), WrapContents = false };
            _txtNewTask = new TextBox { Width = 320, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(0, 0, 15, 0) };
            _txtNewTask.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await AddNewTask(); };
            _chkNewTaskHasDate = new CheckBox { Text = "期限:", AutoSize = true, Cursor = Cursors.Hand, Margin = new Padding(0, 4, 5, 0) };
            _dtpNewTaskDate = new DateTimePicker { Width = 140, Format = DateTimePickerFormat.Short, Enabled = false, Margin = new Padding(0, 0, 25, 0) };
            _chkNewTaskHasDate.CheckedChanged += (s, e) => _dtpNewTaskDate.Enabled = _chkNewTaskHasDate.Checked;
            _btnAdd = new Button { Text = "+ 新增任務", Width = 110, Height = 30, Margin = new Padding(0, -2, 0, 0), BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnAdd.FlatAppearance.BorderSize = 0;
            _btnAdd.Click += async (s, e) => await AddNewTask();
            pnlRow1.Controls.AddRange(new Control[] { _txtNewTask, _chkNewTaskHasDate, _dtpNewTaskDate, _btnAdd });

            FlowLayoutPanel pnlRow2 = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 45, Padding = new Padding(10, 5, 10, 0), WrapContents = false };
            _btnClearDone = new Button { Text = "🧹 清除已完成", Width = 130, Height = 30, Margin = new Padding(0, 0, 20, 0), BackColor = Color.FromArgb(204, 51, 51), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnClearDone.FlatAppearance.BorderSize = 0;
            _btnClearDone.Click += async (s, e) => await ClearAllCompletedTasks();
            _btnExportPdf = new Button { Text = "📄 轉存 PDF", Width = 120, Height = 30, Margin = new Padding(0, 0, 25, 0), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnExportPdf.FlatAppearance.BorderSize = 0;
            _btnExportPdf.Click += (s, e) => ExportToPdf();
            _chkShowCompleted = new CheckBox { Text = "顯示已完成項目", AutoSize = true, Checked = false, Cursor = Cursors.Hand, Margin = new Padding(0, 4, 0, 0) };
            _chkShowCompleted.CheckedChanged += async (s, e) => await SyncTasksSilently();
            pnlRow2.Controls.AddRange(new Control[] { _btnClearDone, _btnExportPdf, _chkShowCompleted });

            pnlTopContainer.Controls.Add(pnlRow2);
            pnlTopContainer.Controls.Add(pnlRow1);

            _taskView = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = false, CheckBoxes = true, BorderStyle = BorderStyle.None, ShowGroups = true, LabelEdit = false };
            _taskView.Columns.Add("刪除", 60, HorizontalAlignment.Center);
            _taskView.Columns.Add("任務名稱 (雙擊進行進階編輯)", 500);
            _taskView.Columns.Add("到期日", 160);
            _taskView.Resize += (s, e) => {
                int contentWidth = _taskView.ClientSize.Width - _taskView.Columns[0].Width - _taskView.Columns[2].Width;
                if (contentWidth > 100) _taskView.Columns[1].Width = contentWidth;
            };

            _groupTodo = new ListViewGroup("📌 待辦事項", HorizontalAlignment.Left);
            _groupDone = new ListViewGroup("✅ 已完成", HorizontalAlignment.Left);
            _taskView.Groups.Add(_groupTodo);
            _taskView.Groups.Add(_groupDone);

            _taskView.ItemCheck += TaskView_ItemCheck; 
            _taskView.ItemChecked += async (s, e) => await HandleTaskChecked(e);
            _taskView.MouseClick += async (s, e) => {
                if (e.Button != MouseButtons.Left) return;
                var hit = _taskView.HitTest(e.Location);
                if (hit.Item != null && hit.Item.SubItems.IndexOf(hit.SubItem) == 0 && hit.Location != ListViewHitTestLocations.StateImage) await DeleteTaskAction(hit.Item);
            };
            _taskView.DoubleClick += async (s, e) => await OpenEditDialog();
            _taskView.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Delete && _taskView.SelectedItems.Count > 0) await DeleteTaskAction(_taskView.SelectedItems[0]); };

            _status = new Label { Dock = DockStyle.Bottom, Height = 25, Text = " 準備就緒", BackColor = Color.WhiteSmoke, ForeColor = Color.DimGray };

            this.Controls.Add(_taskView); this.Controls.Add(pnlTopContainer); this.Controls.Add(_status); this.Controls.Add(menu);
            this.MainMenuStrip = menu;

            this.Shown += async (s, e) => await SyncTasksSilently();
        }

        private void ExportToPdf()
        {
            if (_taskView.Items.Count == 0) return;
            using (PrintDialog printDialog = new PrintDialog())
            {
                PrintDocument printDoc = new PrintDocument();
                printDoc.DocumentName = "G-Task Nexus 待辦清單";
                printDialog.Document = printDoc;
                foreach (string printer in PrinterSettings.InstalledPrinters) {
                    if (printer.Contains("PDF")) { printDoc.PrinterSettings.PrinterName = printer; break; }
                }
                printDoc.PrintPage += PrintDoc_PrintPage;
                if (printDialog.ShowDialog() == DialogResult.OK) {
                    _printGroupIndex = 0; _printItemIndex = 0;
                    printDoc.Print();
                }
            }
        }

        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            float yPos = e.MarginBounds.Top;
            float leftMargin = e.MarginBounds.Left;
            Font titleFont = new Font("Microsoft JhengHei UI", 16, FontStyle.Bold);
            Font groupFont = new Font("Microsoft JhengHei UI", 12, FontStyle.Bold);
            Font itemFont = new Font("Microsoft JhengHei UI", 11, FontStyle.Regular);

            if (_printGroupIndex == 0 && _printItemIndex == 0) {
                e.Graphics.DrawString($"G-Task Nexus 待辦清單 (匯出時間: {DateTime.Now:yyyy/MM/dd})", titleFont, Brushes.Black, leftMargin, yPos);
                yPos += 45;
            }

            while (_printGroupIndex < _taskView.Groups.Count) {
                ListViewGroup group = _taskView.Groups[_printGroupIndex];
                if (_printItemIndex == 0) { e.Graphics.DrawString(group.Header, groupFont, Brushes.DarkBlue, leftMargin, yPos); yPos += 30; }
                while (_printItemIndex < group.Items.Count) {
                    ListViewItem item = group.Items[_printItemIndex];
                    string status = item.Checked ? "☑" : "☐";
                    string title = item.SubItems.Count > 1 ? item.SubItems[1].Text : "";
                    string date = item.SubItems.Count > 2 ? item.SubItems[2].Text : "";
                    e.Graphics.DrawString($"{status} {title}{(string.IsNullOrEmpty(date) ? "" : $" (到期: {date})")}", itemFont, item.Checked ? Brushes.Gray : Brushes.Black, leftMargin + 10, yPos);
                    yPos += 28; _printItemIndex++;
                    if (yPos > e.MarginBounds.Bottom) { e.HasMorePages = true; return; }
                }
                _printItemIndex = 0; _printGroupIndex++; yPos += 15;
            }
            e.HasMorePages = false;
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
            if (MessageBox.Show($"確定要刪除「{(item.SubItems.Count > 1 ? item.SubItems[1].Text : "")}」嗎？", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try { await _service.DeleteTaskAsync(taskId); _taskView.Items.Remove(item); } catch { }
            }
        }

        private async Task ClearAllCompletedTasks()
        {
            if (MessageBox.Show("確定要永久刪除所有「已完成」任務嗎？", "一鍵清除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try { await _service.ClearCompletedTasksAsync(); await SyncTasksSilently(); } catch { }
            }
        }

        private async Task OpenEditDialog()
        {
            var hit = _taskView.HitTest(_taskView.PointToClient(Cursor.Position));
            if (hit.Item == null || hit.Item.SubItems.IndexOf(hit.SubItem) == 0 || _taskView.SelectedItems.Count == 0) return;
            var item = _taskView.SelectedItems[0];
            string taskId = item.Tag?.ToString();
            DateTime? existingDate = null;
            if (item.SubItems.Count > 2 && DateTime.TryParse(item.SubItems[2].Text, out DateTime parsed)) existingDate = parsed;

            using (Form editForm = new Form { Text = "編輯任務詳情", Size = new Size(380, 250), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false }) {
                Label lblTitle = new Label { Text = "任務名稱:", Location = new Point(15, 15), AutoSize = true };
                TextBox txtTitle = new TextBox { Text = item.SubItems[1].Text, Location = new Point(15, 40), Width = 330 };
                CheckBox chkHasDate = new CheckBox { Text = "啟用日期", Location = new Point(15, 105), AutoSize = true, Checked = existingDate.HasValue };
                DateTimePicker dtpDate = new DateTimePicker { Location = new Point(100, 102), Width = 245, Format = DateTimePickerFormat.Short, Enabled = chkHasDate.Checked };
                if (existingDate.HasValue) dtpDate.Value = existingDate.Value;
                chkHasDate.CheckedChanged += (s, e) => dtpDate.Enabled = chkHasDate.Checked;
                Button btnSave = new Button { Text = "儲存變更", Location = new Point(130, 160), Width = 120, Height = 35, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                btnSave.Click += async (s, e) => {
                    _autoSyncTimer.Stop(); _isSyncing = true;
                    try { await _service.UpdateTaskDetailsAsync(taskId, txtTitle.Text, chkHasDate.Checked ? dtpDate.Value : (DateTime?)null); editForm.DialogResult = DialogResult.OK; } 
                    finally { _isSyncing = false; _autoSyncTimer.Start(); }
                };
                editForm.AcceptButton = btnSave;
                editForm.Controls.AddRange(new Control[] { lblTitle, txtTitle, chkHasDate, dtpDate, btnSave });
                if (editForm.ShowDialog() == DialogResult.OK) await SyncTasksSilently();
            }
        }

        private void SetupAutoSync() { _autoSyncTimer = new Timer { Interval = 15000 }; _autoSyncTimer.Tick += async (s, e) => await SyncTasksSilently(); _autoSyncTimer.Start(); }

        private async Task HandleTaskChecked(ItemCheckedEventArgs e)
        {
            if (_isSyncing) return;
            string taskId = e.Item.Tag?.ToString();
            bool isDone = e.Item.Checked;
            if (!string.IsNullOrEmpty(taskId)) {
                e.Item.UseItemStyleForSubItems = false; e.Item.ForeColor = Color.Red; 
                Color textColor = isDone ? Color.Gray : Color.Black;
                Font textFont = new Font(_taskView.Font, isDone ? FontStyle.Strikeout : FontStyle.Regular);
                if (e.Item.SubItems.Count > 1) { e.Item.SubItems[1].ForeColor = textColor; e.Item.SubItems[1].Font = textFont; }
                if (e.Item.SubItems.Count > 2) { e.Item.SubItems[2].ForeColor = textColor; e.Item.SubItems[2].Font = textFont; }
                e.Item.Group = isDone ? _groupDone : _groupTodo;
                try { await _service.UpdateTaskStatusAsync(taskId, isDone); } catch { }
            }
        }

        private async Task AddNewTask()
        {
            string title = _txtNewTask.Text.Trim();
            if (string.IsNullOrEmpty(title)) return;
            try { await _service.AddTaskAsync(title, _chkNewTaskHasDate.Checked ? _dtpNewTaskDate.Value : (DateTime?)null); _txtNewTask.Clear(); _chkNewTaskHasDate.Checked = false; await SyncTasksSilently(); } catch { }
        }

        private async Task SyncTasksSilently()
        {
            if (_isSyncing) return; _isSyncing = true;
            try {
                var items = await _service.GetAllTasksAsync(_chkShowCompleted.Checked);
                _taskView.BeginUpdate(); _taskView.Items.Clear();
                foreach (var t in items) {
                    var lvi = new ListViewItem(" ❌") { Tag = t.Id, Checked = (t.Status == "completed"), UseItemStyleForSubItems = false, ForeColor = Color.Red };
                    var subTitle = lvi.SubItems.Add(t.Title ?? "無標題");
                    string dateStr = "";
                    if (!string.IsNullOrEmpty(t.Due) && DateTime.TryParse(t.Due, out DateTime parsedDate)) dateStr = parsedDate.ToString("yyyy/MM/dd");
                    var subDate = lvi.SubItems.Add(dateStr);
                    Color rowColor = lvi.Checked ? Color.Gray : Color.Black;
                    Font rowFont = new Font(_taskView.Font, lvi.Checked ? FontStyle.Strikeout : FontStyle.Regular);
                    subTitle.ForeColor = rowColor; subTitle.Font = rowFont; subDate.ForeColor = rowColor; subDate.Font = rowFont;
                    lvi.Group = lvi.Checked ? _groupDone : _groupTodo;
                    _taskView.Items.Add(lvi);
                }
                _taskView.EndUpdate();
                UpdateStatus($" 同步完成。最後更新: {DateTime.Now:HH:mm:ss}");
            } catch { } finally { _isSyncing = false; }
        }

        private void OpenApiConfig()
        {
            using (Form f = new Form { Text = "貼入 API JSON", Size = new Size(400, 300), StartPosition = FormStartPosition.CenterParent }) {
                TextBox txt = new TextBox { Multiline = true, Dock = DockStyle.Fill, Text = CoreSecurity.LoadSecureData() };
                Button btn = new Button { Text = "儲存並重新啟動", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.FromArgb(45,45,45), ForeColor=Color.White };
                btn.Click += (s, e) => { CoreSecurity.SaveSecureData(txt.Text); Application.Restart(); };
                f.Controls.Add(txt); f.Controls.Add(btn); f.ShowDialog();
            }
        }

        private void UpdateStatus(string msg) { if (this.InvokeRequired) this.Invoke(new Action(() => _status.Text = msg)); else _status.Text = msg; }
    }
}
