/// FILE: Safety_System/settings/App_ViewPermissionManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ViewPermissionManager : Form
    {
        private ListBox _lstUsers;
        private TreeView _tvMenus;
        private Button _btnSave;
        private MenuStrip _mainMenuRef;

        private const string DbName = "SystemConfig";
        private const string TableName = "HiddenUserMenus";

        public App_ViewPermissionManager(MenuStrip mainMenu)
        {
            _mainMenuRef = mainMenu;
            
            // 初始化隱藏名單資料表
            string createSql = "CREATE TABLE IF NOT EXISTS [HiddenUserMenus] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [UserName] TEXT, [MenuText] TEXT, UNIQUE(UserName, MenuText));";
            DataManager.InitTable(DbName, TableName, createSql);

            InitializeComponent();
            LoadUsers();
            BuildMenuTree();
        }

        private void InitializeComponent()
        {
            this.Text = "選單閱覽權限設定 (Lv3 專屬)";
            this.Size = new Size(800, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label { Text = "👁️ 使用者選單閱覽權限設定", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Location = new Point(20, 20), AutoSize = true };
            Label lblDesc = new Label { Text = "操作說明：在左側選擇使用者，右側打勾表示【可見】，取消打勾表示【隱藏】。", Font = new Font("Microsoft JhengHei UI", 11F), ForeColor = Color.DimGray, Location = new Point(25, 55), AutoSize = true };

            SplitContainer split = new SplitContainer { 
                Location = new Point(25, 90), 
                Size = new Size(730, 430), 
                SplitterDistance = 250,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 左側：使用者清單
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            Label lblUser = new Label { Text = "1. 選擇使用者：", Dock = DockStyle.Top, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Height = 30 };
            _lstUsers = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lstUsers.SelectedIndexChanged += LstUsers_SelectedIndexChanged;
            pnlLeft.Controls.Add(_lstUsers);
            pnlLeft.Controls.Add(lblUser);

            // 右側：選單樹狀圖
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            Label lblMenu = new Label { Text = "2. 設定可見選單：", Dock = DockStyle.Top, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Height = 30 };
            _tvMenus = new TreeView { 
                Dock = DockStyle.Fill, 
                CheckBoxes = true, 
                Font = new Font("Microsoft JhengHei UI", 12F),
                ItemHeight = 28
            };
            _tvMenus.AfterCheck += TvMenus_AfterCheck;
            pnlRight.Controls.Add(_tvMenus);
            pnlRight.Controls.Add(lblMenu);

            split.Panel1.Controls.Add(pnlLeft);
            split.Panel2.Controls.Add(pnlRight);

            _btnSave = new Button { 
                Text = "💾 儲存該使用者的權限設定", 
                Location = new Point(25, 540), 
                Size = new Size(730, 45), 
                BackColor = Color.ForestGreen, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                Cursor = Cursors.Hand 
            };
            _btnSave.Click += BtnSave_Click;

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblDesc);
            this.Controls.Add(split);
            this.Controls.Add(_btnSave);
        }

        private void LoadUsers()
        {
            _lstUsers.Items.Clear();
            
            // 加入預設的高級權限使用者
            foreach (var u in LicenseManager.DefaultUsers) {
                if (!_lstUsers.Items.Contains(u)) _lstUsers.Items.Add(u);
            }

            // 加入系統自訂授權的使用者
            try {
                DataTable dt = DataManager.GetTableData("SystemConfig", "AllowedUsers", "", "", "");
                if (dt != null) {
                    foreach (DataRow row in dt.Rows) {
                        string u = row["使用者帳號"].ToString().Trim();
                        if (!string.IsNullOrEmpty(u) && !_lstUsers.Items.Contains(u)) {
                            _lstUsers.Items.Add(u);
                        }
                    }
                }
            } catch { }

            if (_lstUsers.Items.Count > 0) _lstUsers.SelectedIndex = 0;
        }

        // 動態抓取 MainForm 的選單生成樹狀圖
        private void BuildMenuTree()
        {
            _tvMenus.Nodes.Clear();
            if (_mainMenuRef == null) return;

            foreach (ToolStripItem item in _mainMenuRef.Items)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    // 忽略被隱藏的特殊選單 (如選單1~4)
                    if (!menuItem.Visible && menuItem.Text.StartsWith("選單")) continue;

                    TreeNode node = new TreeNode(menuItem.Text);
                    node.Checked = true;
                    AddChildNodes(node, menuItem);
                    _tvMenus.Nodes.Add(node);
                }
            }
            _tvMenus.ExpandAll();
        }

        private void AddChildNodes(TreeNode parentNode, ToolStripMenuItem parentMenuItem)
        {
            foreach (ToolStripItem item in parentMenuItem.DropDownItems)
            {
                if (item is ToolStripMenuItem subMenu)
                {
                    // 忽略「開啟個人隱藏選單」、「變更個人選單密碼」、「記憶體釋放」等系統工具
                    if (subMenu.Text.Contains("個人隱藏") || subMenu.Text.Contains("記憶體釋放")) continue;

                    TreeNode childNode = new TreeNode(subMenu.Text);
                    childNode.Checked = true;
                    AddChildNodes(childNode, subMenu);
                    parentNode.Nodes.Add(childNode);
                }
            }
        }

        // 父節點與子節點勾選連動邏輯
        private bool _isUpdatingCheck = false;
        private void TvMenus_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_isUpdatingCheck) return;
            _isUpdatingCheck = true;

            // 往下連動：勾選/取消父節點，子節點同步
            CheckAllChildNodes(e.Node, e.Node.Checked);
            
            // 往上連動：如果子節點打勾，父節點一定要打勾
            if (e.Node.Checked && e.Node.Parent != null) {
                CheckParentNode(e.Node.Parent);
            }

            _isUpdatingCheck = false;
        }

        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0) CheckAllChildNodes(node, nodeChecked);
            }
        }

        private void CheckParentNode(TreeNode treeNode)
        {
            treeNode.Checked = true;
            if (treeNode.Parent != null) CheckParentNode(treeNode.Parent);
        }

        private void LstUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lstUsers.SelectedItem == null) return;
            string targetUser = _lstUsers.SelectedItem.ToString();

            // 1. 先將所有選單設為打勾 (可見)
            _isUpdatingCheck = true;
            foreach (TreeNode node in _tvMenus.Nodes) {
                SetNodeCheckState(node, true);
            }
            _isUpdatingCheck = false;

            // 2. 從資料庫讀取該使用者的「黑名單」，並將其取消打勾
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT MenuText FROM HiddenUserMenus WHERE UserName=@U", conn)) {
                        cmd.Parameters.AddWithValue("@U", targetUser);
                        using (var reader = cmd.ExecuteReader()) {
                            while (reader.Read()) {
                                string hiddenMenu = reader["MenuText"].ToString();
                                UncheckNodeByText(_tvMenus.Nodes, hiddenMenu);
                            }
                        }
                    }
                }
            } catch { }
        }

        private void SetNodeCheckState(TreeNode node, bool state)
        {
            node.Checked = state;
            foreach (TreeNode child in node.Nodes) SetNodeCheckState(child, state);
        }

        private void UncheckNodeByText(TreeNodeCollection nodes, string searchText)
        {
            foreach (TreeNode node in nodes) {
                if (node.Text == searchText) {
                    node.Checked = false;
                    // 若父節點被取消打勾，子節點視覺上也應一併取消
                    SetNodeCheckState(node, false); 
                }
                UncheckNodeByText(node.Nodes, searchText);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_lstUsers.SelectedItem == null) return;
            string targetUser = _lstUsers.SelectedItem.ToString();

            // 找出所有「沒有被打勾」的節點文字 (這些是要隱藏的)
            List<string> hiddenMenus = new List<string>();
            GetUncheckedNodes(_tvMenus.Nodes, hiddenMenus);

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        
                        // 先清空該使用者舊的設定
                        using (var cmdDel = new SQLiteCommand("DELETE FROM HiddenUserMenus WHERE UserName=@U", conn, trans)) {
                            cmdDel.Parameters.AddWithValue("@U", targetUser);
                            cmdDel.ExecuteNonQuery();
                        }

                        // 寫入新的隱藏名單
                        string sqlInsert = "INSERT INTO HiddenUserMenus (UserName, MenuText) VALUES (@U, @M)";
                        using (var cmdIns = new SQLiteCommand(sqlInsert, conn, trans)) {
                            cmdIns.Parameters.AddWithValue("@U", targetUser);
                            foreach (var menu in hiddenMenus) {
                                cmdIns.Parameters.AddWithValue("@M", menu);
                                cmdIns.ExecuteNonQuery();
                            }
                        }

                        trans.Commit();
                    }
                }
                MessageBox.Show($"使用者【{targetUser}】的閱覽權限設定已儲存！\n\n(該使用者下次重新開啟系統時生效)", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                MessageBox.Show("儲存失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetUncheckedNodes(TreeNodeCollection nodes, List<string> list)
        {
            foreach (TreeNode node in nodes) {
                if (!node.Checked) {
                    list.Add(node.Text);
                }
                GetUncheckedNodes(node.Nodes, list);
            }
        }
    }
}
