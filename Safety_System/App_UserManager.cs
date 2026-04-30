/// FILE: Safety_System/App_UserManager.cs ///
using System;
using System.Data;
using System.Drawing;
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
            this.Size = new Size(500, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label { Text = "👤 電腦登入帳號授權清單", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.SteelBlue, Location = new Point(25, 20), AutoSize = true };
            Label lblDesc = new Label { Text = "※ 此處的帳號必須與使用者登入 Windows 的電腦帳戶名稱完全一致。", Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.DimGray, Location = new Point(30, 55), AutoSize = true };

            GroupBox boxAdd = new GroupBox { Text = "新增授權帳號", Location = new Point(25, 90), Size = new Size(435, 100), Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtNewUser = new TextBox { Location = new Point(20, 40), Width = 250 };
            
            Button btnAdd = new Button { Text = "➕ 新增", Location = new Point(290, 38), Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnAdd.FlatAppearance.BorderSize = 0;
            btnAdd.Click += BtnAdd_Click;
            
            boxAdd.Controls.Add(_txtNewUser);
            boxAdd.Controls.Add(btnAdd);

            GroupBox boxList = new GroupBox { Text = "已授權之帳號 (點擊刪除)", Location = new Point(25, 210), Size = new Size(435, 320), Font = new Font("Microsoft JhengHei UI", 12F) };
            _flpUsers = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxList.Controls.Add(_flpUsers);

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblDesc);
            this.Controls.Add(boxAdd);
            this.Controls.Add(boxList);
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

                Panel pnl = new Panel { Width = 390, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                Label lbl = new Label { Text = user, Location = new Point(15, 12), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                
                Button btnDel = new Button { Text = "🗑️ 刪除", Width = 80, Height = 35, Location = new Point(300, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
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

            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            foreach (DataRow row in dt.Rows)
            {
                if (row["使用者帳號"].ToString().Equals(newUser, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("該帳號已經存在於授權清單中！", "重複", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show($"帳號【{newUser}】已成功加入授權名單！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
