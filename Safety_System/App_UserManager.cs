/// FILE: Safety_System/App_UserManager.cs ///
using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_UserManager : Form
    {
        private TextBox _txtNewUser;
        private FlowLayoutPanel _flpUsers;
        private const string DbName = "SystemConfig";
        private const string TableName = "AllowedUsers";

        public App_UserManager()
        {
            InitializeComponent();
            LoadUsers();
        }

        private void InitializeComponent()
        {
            this.Text = "軟體啟用認證 - 授權帳號管理";
            // 🟢 調整 1：將視窗寬度從 500 加寬至 550，給標題和按鈕足夠的空間
            this.Size = new Size(550, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label { Text = "👤 電腦登入帳號授權清單", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.SteelBlue, Location = new Point(25, 20), AutoSize = true };
            Label lblDesc = new Label { Text = "※ 此處的帳號必須與使用者登入 Windows 的電腦帳戶名稱完全一致。", Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.DimGray, Location = new Point(30, 55), AutoSize = true };

            // 🟢 調整 2：右上角按鈕往右移動 (X 從 275 改為 325)
            Button btnShowDefault = new Button { 
                Text = "🛡️ 檢視系統預設帳號", 
                Location = new Point(325, 20), 
                Size = new Size(185, 35), 
                BackColor = Color.SlateGray, 
                ForeColor = Color.White, 
                Cursor = Cursors.Hand, 
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold)
            };
            btnShowDefault.FlatAppearance.BorderSize = 0;
            btnShowDefault.Click += BtnShowDefault_Click;

            // 🟢 調整 3：跟著視窗加寬，外框寬度從 435 增加到 485
            GroupBox boxAdd = new GroupBox { Text = "新增自訂授權帳號", Location = new Point(25, 90), Size = new Size(485, 100), Font = new Font("Microsoft JhengHei UI", 12F) };
            
            // 🟢 調整 4：輸入框加寬
            _txtNewUser = new TextBox { Location = new Point(20, 40), Width = 300 };
            
            // 🟢 調整 5：新增按鈕跟著往右移
            Button btnAdd = new Button { Text = "➕ 新增", Location = new Point(340, 38), Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnAdd.FlatAppearance.BorderSize = 0;
            btnAdd.Click += BtnAdd_Click;
            
            boxAdd.Controls.Add(_txtNewUser);
            boxAdd.Controls.Add(btnAdd);

            // 🟢 調整 6：清單外框也跟著加寬
            GroupBox boxList = new GroupBox { Text = "已自訂授權之帳號 (點擊刪除)", Location = new Point(25, 210), Size = new Size(485, 320), Font = new Font("Microsoft JhengHei UI", 12F) };
            _flpUsers = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxList.Controls.Add(_flpUsers);

            this.Controls.Add(lblTitle);
            this.Controls.Add(btnShowDefault);
            this.Controls.Add(lblDesc);
            this.Controls.Add(boxAdd);
            this.Controls.Add(boxList);
        }

        private void BtnShowDefault_Click(object sender, EventArgs e)
        {
            string defaultList = string.Join("\n", LicenseManager.DefaultUsers.Select(u => "• " + u));
            
            MessageBox.Show($"以下為系統內建的最高權限預設帳號，擁有永久開啟權限，不可被刪除：\n\n{defaultList}", 
                            "系統預設帳號清單", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void LoadUsers()
        {
            _flpUsers.Controls.Clear();
            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            if (dt == null || dt.Rows.Count == 0) return;

            foreach (DataRow row in dt.Rows)
            {
                int id = Convert.ToInt32(row["Id"]);
                string user = row["使用者帳號"].ToString();

                bool isDefault = LicenseManager.DefaultUsers.Contains(user, StringComparer.OrdinalIgnoreCase);
                if (isDefault) continue;

                // 🟢 調整 7：清單內項目的背景板加寬 (從 390 -> 440)
                Panel pnl = new Panel { Width = 440, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                Label lbl = new Label { Text = user, Location = new Point(15, 12), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                
                // 🟢 調整 8：刪除按鈕跟著往右移 (從 300 -> 350)
                Button btnDel = new Button { Text = "🗑️ 刪除", Width = 80, Height = 35, Location = new Point(350, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                btnDel.FlatAppearance.BorderSize = 0;
                btnDel.Click += (s, e) => {
                    if (MessageBox.Show($"確定要移除【{user}】的軟體開啟權限嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        DataManager.DeleteRecord(DbName, TableName, id);
                        LoadUsers();
                    }
                };

                pnl.Controls.Add(lbl);
                pnl.Controls.Add(btnDel);
                _flpUsers.Controls.Add(pnl);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string newUser = _txtNewUser.Text.Trim();
            if (string.IsNullOrEmpty(newUser)) return;

            if (LicenseManager.DefaultUsers.Contains(newUser, StringComparer.OrdinalIgnoreCase))
            {
                MessageBox.Show("該帳號屬於「系統預設帳號」，已具備永久權限，無須重複新增！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _txtNewUser.Clear();
                return;
            }

            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            foreach (DataRow row in dt.Rows)
            {
                if (row["使用者帳號"].ToString().Equals(newUser, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("該帳號已經存在於自訂授權清單中！", "重複", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            DataRow newRow = dt.NewRow();
            newRow["使用者帳號"] = newUser;
            dt.Rows.Add(newRow);
            
            if (DataManager.BulkSaveTable(DbName, TableName, dt))
            {
                _txtNewUser.Clear();
                LoadUsers();
                MessageBox.Show($"帳號【{newUser}】已成功加入自訂授權名單！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
