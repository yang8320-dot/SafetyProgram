/// FILE: Safety_System/settings/App_PermissionManager.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_PermissionManager : Form
    {
        private MenuStrip _mainMenuRef;

        public App_PermissionManager(MenuStrip mainMenu)
        {
            _mainMenuRef = mainMenu;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "系統權限設定總管 (Lv3 專屬)";
            this.Size = new Size(850, 720); // 確保足夠寬廣以容納兩個子頁面
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.WhiteSmoke;

            TabControl tabControl = new TabControl {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft JhengHei UI", 12F),
                Padding = new Point(15, 8)
            };

            // 建立兩個頁籤
            TabPage tabUser = new TabPage("👤 電腦登入帳號授權清單");
            tabUser.BackColor = Color.White;
            
            TabPage tabView = new TabPage("👁️ 選單閱覽權限設定");
            tabView.BackColor = Color.White;

            // 將原本的 App_UserManager 嵌入 Tab 1
            App_UserManager frmUser = new App_UserManager();
            frmUser.TopLevel = false;
            frmUser.FormBorderStyle = FormBorderStyle.None;
            frmUser.Dock = DockStyle.Fill;
            tabUser.Controls.Add(frmUser);
            frmUser.Show();

            // 將原本的 App_ViewPermissionManager 嵌入 Tab 2
            App_ViewPermissionManager frmView = new App_ViewPermissionManager(_mainMenuRef);
            frmView.TopLevel = false;
            frmView.FormBorderStyle = FormBorderStyle.None;
            frmView.Dock = DockStyle.Fill;
            tabView.Controls.Add(frmView);
            frmView.Show();

            tabControl.TabPages.Add(tabUser);
            tabControl.TabPages.Add(tabView);

            this.Controls.Add(tabControl);
        }
    }
}
