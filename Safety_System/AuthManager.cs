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
        // 🟢 定義兩個等級的密碼
        private const string UserPassword = "1234";        // 一般權限
        private const string AdminPassword = "11914002";   // 管理者權限

        /// <summary>
        /// 驗證一般使用者權限 (管理員密碼亦可通過)
        /// 用於：新增/修改欄位、刪除紀錄等日常維護操作
        /// </summary>
        public static bool VerifyUser(string prompt = "請輸入操作授權密碼：")
        {
            string input = ShowAuthDialog(prompt);
            // 輸入一般密碼或管理員密碼皆可放行
            return input == UserPassword || input == AdminPassword;
        }

        /// <summary>
        /// 驗證管理者權限 (僅管理員密碼可通過)
        /// 用於：資料庫路徑設定、防重寫規則設定等核心系統操作
        /// </summary>
        public static bool VerifyAdmin(string prompt = "此為系統核心設定，請輸入【管理者】密碼：")
        {
            string input = ShowAuthDialog(prompt);
            // 僅限管理員密碼可放行
            return input == AdminPassword;
        }

        // 🟢 私有對話框邏輯，供內部共用
        private static string ShowAuthDialog(string prompt)
        {
            using (Form p = new Form())
            {
                p.Width = 500; 
                p.Height = 300;
                p.Text = "權限驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;

                Label lbl = new Label() { Left = 30, Top = 30, Text = prompt, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 370, Left = 30, Top = 80, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認", DialogResult = DialogResult.OK, Left = 280, Top = 150, Width = 120, Height = 40, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog(Form.ActiveForm) == DialogResult.OK) return txt.Text;
                return "";
            }
        }
    }
}
