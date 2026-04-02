using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class MainForm : Form
    {
        private MenuStrip _mainMenu;
        private Panel _contentPanel;

        public MainForm()
        {
            this.Text = "工安系統看板 (v4.7.2)";
            this.Size = new Size(1440, 810);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1440, 810);

            _mainMenu = new MenuStrip();
            BuildMenu();
            this.Controls.Add(_mainMenu);

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(30),
                AutoScroll = true
            };
            this.Controls.Add(_contentPanel);
            
            _mainMenu.BringToFront();
            LoadWelcomeScreen();
        }

        // ==========================================
        // 1. 新增快捷鍵 Ctrl + 3 攔截邏輯
        // ==========================================
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.D3) || keyData == (Keys.Control | Keys.NumPad3))
            {
                // TODO: 這裡你可以改為你要的功能，例如：
                // LoadModule(new App_SafetyInspection().GetView());
                
                MessageBox.Show("觸發快捷鍵：Ctrl + 3\n(可在此設定快速開啟某個模組)", "快捷鍵", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                return true; // 代表按鍵事件已處理，不要再往下傳遞
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void BuildMenu()
        {
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            
            // 待辦清單 (若你剛剛有實作 App_TodoList.cs，可取消註解下面這兩行)
            // var itemTodo = new ToolStripMenuItem("待辦清單");
            // itemTodo.Click += (s, e) => LoadModule(new App_TodoList().GetView());

            menuSafety.DropDownItems.Add(itemInspection);
            // menuSafety.DropDownItems.Add(itemTodo);
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("虛驚"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("工傷"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("交傷"));

            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(new ToolStripMenuItem("空污申報"));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(new ToolStripMenuItem("消防設備"));

            // ==========================================
            // 2. 新增設定選單 -> 設定資料庫
            // ==========================================
            var menuSettings = new ToolStripMenuItem("設定");
            var itemDbConfig = new ToolStripMenuItem("設定資料庫");
            itemDbConfig.Click += ItemDbConfig_Click;
            menuSettings.DropDownItems.Add(itemDbConfig);

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), 
                new ToolStripMenuItem("溫盤"), 
                menuSettings
            });
        }

        // ==========================================
        // 3. 設定資料庫路徑的彈出視窗事件
        // ==========================================
        private void ItemDbConfig_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "請選擇「工安系統資料庫 (純文字檔)」要集中儲存的資料夾：";
                fbd.SelectedPath = DataManager.BasePath; // 預設指向當前設定的路徑
                
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    DataManager.SetBasePath(fbd.SelectedPath);
                    MessageBox.Show($"系統資料儲存路徑已更新為：\n{fbd.SelectedPath}", "設定成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        public void LoadModule(Control moduleControl)
        {
            _contentPanel.Controls.Clear();
            moduleControl.Dock = DockStyle.Fill;
            _contentPanel.Controls.Add(moduleControl);
        }

        private void LoadWelcomeScreen()
        {
            Label lbl = new Label
            {
                Text = "=== 工安系統看板 ===\n架構：.NET Framework 4.7.2",
                Font = new Font("Microsoft JhengHei", 24, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(50, 50)
            };
            _contentPanel.Controls.Add(lbl);
        }
    }
}
