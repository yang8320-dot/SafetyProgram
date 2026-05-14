/// FILE: Safety_System/App_PasswordManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_PasswordManager : Form
    {
        private const string DbName = "SystemConfig";
        private const string TableName = "MenuPasswords";

        private ComboBox _cboMenu;
        private TextBox _txtOldPwd;
        private TextBox _txtNewPwd;
        private TextBox _txtHint;

        // 預設密碼字典
        public static readonly Dictionary<string, string> DefaultPasswords = new Dictionary<string, string>
        {
            { "選單1", "7361" },
            { "選單2", "7362" },
            { "選單3", "7363" },
            { "選單4", "7328" }
        };

        public App_PasswordManager()
        {
            InitDatabase();
            InitializeComponent();
        }

        // 初始化密碼管理表
        public static void InitDatabase()
        {
            string schema = "[選單名稱] TEXT, [密碼] TEXT, [提示詞] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{TableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(DbName, TableName, createSql);
        }

        // 驗證密碼邏輯
        public static string CheckUnlockMenu(string inputPwd)
        {
            InitDatabase();
            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");

            // 1. 先比對資料庫內使用者自訂的密碼
            foreach (DataRow row in dt.Rows)
            {
                if (row["密碼"].ToString() == inputPwd)
                {
                    return row["選單名稱"].ToString();
                }
            }

            // 2. 若資料庫無符合，則比對預設密碼
            foreach (var kvp in DefaultPasswords)
            {
                if (kvp.Value == inputPwd)
                {
                    bool isCustomized = false;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["選單名稱"].ToString() == kvp.Key)
                        {
                            isCustomized = true;
                            break;
                        }
                    }

                    if (!isCustomized) return kvp.Key;
                }
            }

            return null; // 密碼錯誤
        }

        private void InitializeComponent()
        {
            this.Text = "個人選單密碼管理";
            this.Size = new Size(500, 560); // 稍微加高視窗，容納自動換行的備註
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.WhiteSmoke; // 底色調整，增加質感

            Label lblTitle = new Label { Text = "🔐 變更個人選單密碼與提示詞", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Location = new Point(40, 20), AutoSize = true };

            int labelX = 40;
            int inputX = 150;
            int inputW = 280;

            Label lbl1 = new Label { Text = "選擇選單：", Location = new Point(labelX, 85), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboMenu = new ComboBox { Location = new Point(inputX, 82), Width = inputW, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 13F) };
            _cboMenu.Items.AddRange(new string[] { "選單1", "選單2", "選單3", "選單4" });
            _cboMenu.SelectedIndex = 0;
            _cboMenu.SelectedIndexChanged += CboMenu_SelectedIndexChanged;

            Label lblOld = new Label { Text = "原密碼：", Location = new Point(labelX, 145), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtOldPwd = new TextBox { Location = new Point(inputX, 142), Width = inputW, Font = new Font("Microsoft JhengHei UI", 13F), PasswordChar = '*' };

            Label lblNew = new Label { Text = "新密碼：", Location = new Point(labelX, 205), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtNewPwd = new TextBox { Location = new Point(inputX, 202), Width = inputW, Font = new Font("Microsoft JhengHei UI", 13F), PasswordChar = '*' };

            Label lblHint = new Label { Text = "提示詞：", Location = new Point(labelX, 265), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtHint = new TextBox { Location = new Point(inputX, 262), Width = inputW, Font = new Font("Microsoft JhengHei UI", 13F) };

            // 🟢 將 AutoSize 設為 false，並指定 Width 和 Height，文字過長會自動折行
            Label lblDesc = new Label { 
                Text = "※ 忘記密碼：請在上方輸入「提示詞」並點擊忘記密碼以查詢。\n※ 特權變更：若您輸入的是系統管理者密碼(Lv3)，可忽略原密碼限制直接覆蓋。", 
                Location = new Point(labelX, 320), 
                AutoSize = false, 
                Width = 400,
                Height = 60,
                ForeColor = Color.DimGray, 
                Font = new Font("Microsoft JhengHei UI", 10F) 
            };

            Button btnSave = new Button { Text = "💾 變更密碼", Location = new Point(40, 410), Size = new Size(185, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnSave.Click += BtnSave_Click;

            Button btnForgot = new Button { Text = "❓ 忘記密碼", Location = new Point(245, 410), Size = new Size(185, 45), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnForgot.Click += BtnForgot_Click;

            this.Controls.Add(lblTitle);
            this.Controls.Add(lbl1);
            this.Controls.Add(_cboMenu);
            this.Controls.Add(lblOld);
            this.Controls.Add(_txtOldPwd);
            this.Controls.Add(lblNew);
            this.Controls.Add(_txtNewPwd);
            this.Controls.Add(lblHint);
            this.Controls.Add(_txtHint);
            this.Controls.Add(lblDesc);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnForgot);

            ClearInputs();
        }

        private void CboMenu_SelectedIndexChanged(object sender, EventArgs e)
        {
            ClearInputs();
        }

        private void ClearInputs()
        {
            _txtOldPwd.Text = "";
            _txtNewPwd.Text = "";
            _txtHint.Text = "";
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string menuName = _cboMenu.SelectedItem.ToString();
            string currentPwd = DefaultPasswords[menuName]; 
            
            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            DataRow targetRow = null;

            foreach (DataRow row in dt.Rows)
            {
                if (row["選單名稱"].ToString() == menuName)
                {
                    targetRow = row;
                    currentPwd = row["密碼"].ToString();
                    break;
                }
            }

            if (string.IsNullOrWhiteSpace(_txtNewPwd.Text))
            {
                MessageBox.Show("【新密碼】不能為空白！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 🟢 密碼驗證邏輯：如果原密碼輸入的是 Lv3 密碼，允許直接覆蓋 (特權變更)
            if (_txtOldPwd.Text != currentPwd)
            {
                if (_txtOldPwd.Text == "admin") // 🟢 判斷是否為 Lv3
                {
                    MessageBox.Show("已使用【系統管理者權限 Lv3】放行變更！", "特權授權", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("【原密碼】輸入錯誤！請確認後再試。", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // 儲存邏輯
            if (targetRow != null)
            {
                targetRow["密碼"] = _txtNewPwd.Text;
                targetRow["提示詞"] = _txtHint.Text;
            }
            else
            {
                targetRow = dt.NewRow();
                targetRow["選單名稱"] = menuName;
                targetRow["密碼"] = _txtNewPwd.Text;
                targetRow["提示詞"] = _txtHint.Text;
                dt.Rows.Add(targetRow);
            }

            if (DataManager.BulkSaveTable(DbName, TableName, dt))
            {
                MessageBox.Show($"【{menuName}】 密碼變更成功！\n下次請使用新密碼解鎖。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("密碼儲存失敗！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnForgot_Click(object sender, EventArgs e)
        {
            string menuName = _cboMenu.SelectedItem.ToString();
            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            
            DataRow targetRow = null;
            foreach (DataRow row in dt.Rows)
            {
                if (row["選單名稱"].ToString() == menuName)
                {
                    targetRow = row;
                    break;
                }
            }

            if (targetRow == null)
            {
                MessageBox.Show($"【{menuName}】 尚未自訂過密碼，請直接使用系統出廠的預設密碼。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string actualHint = targetRow["提示詞"].ToString();
            string inputHint = _txtHint.Text.Trim();

            if (actualHint == inputHint)
            {
                string actualPwd = targetRow["密碼"].ToString();
                MessageBox.Show($"驗證成功！【{menuName}】 的現有密碼為：\n\n「 {actualPwd} 」", "密碼查詢", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("輸入的【提示詞】錯誤，無法查看密碼！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
