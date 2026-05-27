/// FILE: Safety_System/settings/AuthManager.cs ///
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
        /// </summary>
        public static bool VerifyUser(string prompt = "執行此操作需要系統權限\n請輸入【Lv1一般操作】等級以上\n密碼進行授權：")
        {
            return ShowAuthDialog(prompt, input => input == Pwd_Lv1 || input == Pwd_Lv2 || input == Pwd_Lv3);
        }

        /// <summary>
        /// 驗證管理者權限 (僅 Lv2, Lv3 可通過)
        /// </summary>
        public static bool VerifyAdmin(string prompt = "進入設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：")
        {
            return ShowAuthDialog(prompt, input => input == Pwd_Lv2 || input == Pwd_Lv3);
        }

        /// <summary>
        /// 驗證管理者權限 (嚴格限定僅使用 Lv2 密碼)
        /// </summary>
        public static bool VerifyLv2Only(string prompt = "執行此操作需要系統權限\n請輸入【Lv2管理者】\n密碼進行授權：")
        {
            return ShowAuthDialog(prompt, input => input == Pwd_Lv2);
        }

        /// <summary>
        /// 驗證系統管理者權限 (嚴格限定僅使用 Lv3 密碼)
        /// </summary>
        public static bool VerifyLv3Only(string prompt = "執行此操作需要系統權限\n請輸入【Lv3系統管理者】\n密碼進行授權：")
        {
            return ShowAuthDialog(prompt, input => input == Pwd_Lv3);
        }

        /// <summary>
        /// 驗證刪除資料表權限 (要求依序輸入 Lv2 再輸入 Lv3)
        /// </summary>
        public static bool VerifyTableDelete()
        {
            if (!VerifyLv2Only("此為毀滅性操作(1/2)！\n請先輸入【Lv2管理者】\n密碼進行授權：")) 
            {
                return false;
            }

            if (!VerifyLv3Only("驗證通過(1/2)！\n請接著輸入【Lv3系統管理者】\n密碼以執行最終確認："))
            {
                return false;
            }

            return true;
        }

        // ====================================================================
        // 🟢 全新集中管理：個人隱藏選單專用驗證邏輯
        // ====================================================================
        public static bool VerifyHiddenMenu(string menuName)
        {
            // 1. 檢查是否符合預設擁有者帳號 (免密碼放行)
            string currentUser = Environment.UserName.Trim();
            if (menuName == "選單1" && (string.Equals(currentUser, "黃忠揚", StringComparison.OrdinalIgnoreCase) || string.Equals(currentUser, "TJ700657", StringComparison.OrdinalIgnoreCase))) return true;
            if (menuName == "選單2" && string.Equals(currentUser, "TJ700228", StringComparison.OrdinalIgnoreCase)) return true;
            if (menuName == "選單3" && string.Equals(currentUser, "TJ700533", StringComparison.OrdinalIgnoreCase)) return true;
            if (menuName == "選單4" && string.Equals(currentUser, "TJ204159", StringComparison.OrdinalIgnoreCase)) return true;

            // 2. 若非作者本人，則彈出專屬驗證視窗 (修正排版置中與防遮蔽)
            using (Form p = new Form())
            {
                p.Width = 480;  
                p.Height = 250; 
                p.Text = "個人選單安全驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;
                p.BackColor = Color.White;

                Label lbl = new Label { 
                    Left = 40, Top = 25, AutoSize = true, 
                    Text = "查看此隱藏選單資料表，\n請輸入【" + menuName + "】的解鎖密碼：", 
                    Font = new Font("Microsoft JhengHei UI", 12F) 
                };

                TextBox txt = new TextBox { PasswordChar = '*', Width = 280, Top = 85, Font = new Font("Microsoft JhengHei UI", 14F) };
                txt.Left = (p.ClientSize.Width - txt.Width) / 2; // 水平置中

                Button btn = new Button { 
                    Text = "確認驗證", DialogResult = DialogResult.OK, 
                    Width = 120, Height = 40, Top = 140, 
                    BackColor = Color.SteelBlue, ForeColor = Color.White, 
                    Font = new Font("Microsoft JhengHei UI", 12F) 
                };
                btn.Left = (p.ClientSize.Width - btn.Width) / 2; // 水平置中

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog(Form.ActiveForm) == DialogResult.OK)
                {
                    string input = txt.Text.Trim();
                    string unlockedMenu = App_PasswordManager.CheckUnlockMenu(input);
                    if (unlockedMenu == menuName) return true;
                    
                    MessageBox.Show("【" + menuName + "】密碼錯誤！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false; 
            }
        }

        // ====================================================================
        // 核心視窗邏輯：加入驗證失敗彈窗與重試機制
        // ====================================================================
        private static bool ShowAuthDialog(string prompt, Func<string, bool> validator)
        {
            using (Form p = new Form())
            {
                p.Width = 420; 
                p.Height = 280; 
                p.Text = "權限驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;
                p.BackColor = Color.WhiteSmoke;

                Label lbl = new Label() { Left = 30, Top = 25, Width = 350, Height = 80, Text = prompt, AutoSize = false, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 340, Left = 30, Top = 115, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認授權", Left = 135, Top = 175, Width = 130, Height = 45, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };

                bool isAuthorized = false;

                btn.Click += (s, e) => {
                    if (validator(txt.Text)) {
                        isAuthorized = true;
                        p.DialogResult = DialogResult.OK;
                    } else {
                        MessageBox.Show("密碼錯誤！授權失敗，請重新輸入！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txt.Clear();
                        txt.Focus();
                    }
                };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn; 

                p.ShowDialog(Form.ActiveForm);
                return isAuthorized;
            }
        }
    }
}
