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
            // 設定視窗為 1440 x 810
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

        private void BuildMenu()
        {
            var menuSafety = new ToolStripMenuItem("工安");
            var itemInspection = new ToolStripMenuItem("巡檢");
            itemInspection.Click += (s, e) => LoadModule(new App_SafetyInspection().GetView());
            
            menuSafety.DropDownItems.Add(itemInspection);
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("虛驚"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("工傷"));
            menuSafety.DropDownItems.Add(new ToolStripMenuItem("交傷"));

            var menuEnv = new ToolStripMenuItem("環保");
            menuEnv.DropDownItems.Add(new ToolStripMenuItem("空污申報"));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(new ToolStripMenuItem("消防設備"));

            _mainMenu.Items.AddRange(new ToolStripItem[] {
                menuSafety, menuEnv, menuFire, 
                new ToolStripMenuItem("ESG"), 
                new ToolStripMenuItem("溫盤"), 
                new ToolStripMenuItem("設定")
            });
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
