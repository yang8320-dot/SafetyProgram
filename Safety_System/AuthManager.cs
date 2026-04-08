/// FILE: Safety_System/AuthManager.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 統一權限與密碼驗證管理中心
    /// </summary>
    public static class AuthManager
    {
        // 預設密碼 (未來若要改成從資料庫讀取，只需改這裡)
        private const string AdminPassword = "tces";

        /// <summary>
        /// 彈出密碼驗證視窗
        /// </summary>
        /// <param name="promptMessage">自訂提示文字</param>
        /// <returns>驗證是否成功</returns>
        public static bool VerifyPassword(string promptMessage = "請輸入管理員密碼：")
        {
            using (Form p = new Form())
            {
                p.Width = 450;
                p.Height = 270;
                p.Text = "授權驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false;
                p.MinimizeBox = false;

                Label lbl = new Label() { Left = 30, Top = 30, Text = promptMessage, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                TextBox txtPassword = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btnConfirm = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl);
                p.Controls.Add(txtPassword);
                p.Controls.Add(btnConfirm);
                p.AcceptButton = btnConfirm;

                // 顯示對話框並比對密碼
                return p.ShowDialog(Form.ActiveForm) == DialogResult.OK && txtPassword.Text == AdminPassword;
            }
        }
    }
}
