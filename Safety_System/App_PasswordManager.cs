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
        private TextBox _txtNewPwd;
        private TextBox _txtHint;
        private Label _lblCurrentHint;

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

            // 2. 若資料庫無符合，則比對預設密碼 (前提是該選單還沒被自訂密碼覆蓋)
            foreach (var kvp in DefaultPasswords)
            {
                if (kvp.Value == inputPwd)
                {
                    // 確認資料庫內是否已經有設定過這組選單，若有設定過，預設密碼即失效
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
            this.Size = new Size(450, 480);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label { Text = "🔐 變更個人選單密碼與提示詞", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Location = new Point(30, 20), AutoSize = true };

            Label lbl1 = new Label { Text = "選擇選單：", Location = new Point(30, 80), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboMenu = new ComboBox { Location = new Point(130, 77), Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboMenu.Items.AddRange(new string[] { "選單1", "選單2", "選單3", "選單4" });
            _cboMenu.SelectedIndex = 0;
            _cboMenu.SelectedIndexChanged += CboMenu_SelectedIndexChanged;

            Label lbl2 = new Label { Text = "新密碼：", Location = new Point(30, 140), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtNewPwd = new TextBox { Location = new Point(130, 137), Width = 250, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lbl3 = new Label { Text = "提示詞：", Location = new Point(30, 200), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtHint = new TextBox { Location = new Point(130, 197), Width = 250, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblDesc = new Label { Text = "※ 提示：變更後原預設密碼將失效。\n※ 若忘記密碼，可點擊下方按鈕查詢提示詞。", Location = new Point(30, 260), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 10F) };

            Button btnSave = new Button { Text = "💾 變更密碼", Location = new Point(30, 320), Size = new Size(160, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnSave.Click += BtnSave_Click;

            Button btnForgot = new Button { Text = "❓ 忘記密碼(看提示)", Location = new Point(220, 320), Size = new Size(160, 45), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnForgot.Click += BtnForgot_Click;

            this.Controls.Add(lblTitle);
            this.Controls.Add(lbl1);
            this.Controls.Add(_cboMenu);
            this.Controls.Add(lbl2);
            this.Controls.Add(_txtNewPwd);
            this.Controls.Add(lbl3);
            this.Controls.Add(_txtHint);
            this.Controls.Add(lblDesc);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnForgot);

            LoadCurrentMenuData();
        }

        private void CboMenu_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadCurrentMenuData();
        }

        private void LoadCurrentMenuData()
        {
            string menuName = _cboMenu.SelectedItem.ToString();
            DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
            
            _txtNewPwd.Text = "";
            _txtHint.Text = "";

            foreach (DataRow row in dt.Rows)
            {
                if (row["選單名稱"].ToString() == menuName)
                {
                    _txtNewPwd.Text = row["密碼"].ToString();
                    _txtHint.Text = row["提示詞"].ToString();
                    return;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtNewPwd.Text))
            {
                MessageBox.Show("密碼不能為空白！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
            
            foreach (DataRow row in dt.Rows)
            {
                if (row["選單名稱"].ToString() == menuName)
                {
                    string hint = row["提示詞"].ToString();
                    if (string.IsNullOrWhiteSpace(hint)) hint = "(未設定提示詞)";
                    MessageBox.Show($"【{menuName}】 的密碼提示詞為：\n\n「 {hint} 」", "密碼提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            MessageBox.Show($"【{menuName}】 尚未變更過密碼，請使用系統預設密碼。", "密碼提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
