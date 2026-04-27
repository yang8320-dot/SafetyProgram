/// FILE: Safety_System/AuthManager.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 統一權限與密碼驗證管理中心
    /// 定義權限等級：Lv1(一般), Lv2(管理者), Lv3(系統管理者)
    /// </summary>
    public static class AuthManager
    {
        // 🟢 密碼集中管理
        private const string Pwd_Lv1 = "1234";        // 一般操作密碼
        private const string Pwd_Lv2 = "11914002";    // 管理者密碼
        private const string Pwd_Lv3 = "admin";       // 系統管理者密碼

        /// <summary>
        /// 驗證一般使用者權限 (Lv1, Lv2, Lv3 皆可通過)
        /// 用於：新增/修改欄位、刪除紀錄等日常維護操作
        /// </summary>
        public static bool VerifyUser(string prompt = "請輸入操作授權密碼：")
        {
            string input = ShowAuthDialog(prompt);
            return input == Pwd_Lv1 || input == Pwd_Lv2 || input == Pwd_Lv3;
        }

        /// <summary>
        /// 驗證管理者權限 (僅 Lv2, Lv3 可通過)
        /// 用於：資料庫路徑設定、防重寫規則設定、個人選單密碼變更等
        /// </summary>
        public static bool VerifyAdmin(string prompt = "此為系統核心設定，請輸入【管理者】密碼：")
        {
            string input = ShowAuthDialog(prompt);
            return input == Pwd_Lv2 || input == Pwd_Lv3;
        }

        /// <summary>
        /// 驗證管理者權限 (嚴格限定僅使用 Lv2 密碼)
        /// </summary>
        public static bool VerifyLv2Only(string prompt = "請輸入【管理者】密碼 (Lv2)：")
        {
            string input = ShowAuthDialog(prompt);
            return input == Pwd_Lv2;
        }

        /// <summary>
        /// 驗證系統管理者權限 (嚴格限定僅使用 Lv3 密碼)
        /// </summary>
        public static bool VerifyLv3Only(string prompt = "請輸入【系統管理者】密碼 (Lv3)：")
        {
            string input = ShowAuthDialog(prompt);
            return input == Pwd_Lv3;
        }

        /// <summary>
        /// 🟢 驗證刪除資料表權限 (要求依序輸入 Lv2 再輸入 Lv3)
        /// </summary>
        public static bool VerifyTableDelete()
        {
            if (!VerifyLv2Only("此為毀滅性操作(1/2)！\n請先輸入【管理者密碼 Lv2】：")) 
            {
                MessageBox.Show("Lv2 管理者密碼錯誤，拒絕授權。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!VerifyLv3Only("驗證通過(1/2)！\n請接著輸入【系統管理者密碼 Lv3】以執行最終確認："))
            {
                MessageBox.Show("Lv3 系統管理者密碼錯誤，操作已取消。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
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
