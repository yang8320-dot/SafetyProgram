/// FILE: Safety_System/LicenseManager.cs ///
using System;
using System.Data;
using System.Windows.Forms;

namespace Safety_System
{
    public static class LicenseManager
    {
        private const string DbName = "SystemConfig";
        private const string TableName = "AllowedUsers";
        private static readonly DateTime ExpirationDate = new DateTime(2050, 12, 31);

        // 🟢 將預設帳號名單公開為 public static，讓其他表單可以共用
        public static readonly string[] DefaultUsers = { "黃忠揚", "TJ700657", "TJ700228", "TJ700533", "TJ200248" };

        public static bool VerifyLicense()
        {
            // 1. 檢查軟體使用期限
            if (DateTime.Today > ExpirationDate)
            {
                return false;
            }

            // 2. 初始化使用者資料表
            string createSql = $"CREATE TABLE IF NOT EXISTS [{TableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [使用者帳號] TEXT);";
            DataManager.InitTable(DbName, TableName, createSql);

            // 3. 檢查表內是否有資料，若無則寫入預設名單
            DataTable dtUsers = DataManager.GetTableData(DbName, TableName, "", "", "");
            if (dtUsers == null || dtUsers.Rows.Count == 0)
            {
                DataTable newDt = new DataTable();
                newDt.Columns.Add("使用者帳號", typeof(string));

                foreach (string user in DefaultUsers)
                {
                    newDt.Rows.Add(user);
                }
                DataManager.BulkSaveTable(DbName, TableName, newDt);
                
                // 重新讀取剛寫入的資料
                dtUsers = DataManager.GetTableData(DbName, TableName, "", "", "");
            }

            // 4. 驗證當前電腦登入的帳號 (不分大小寫)
            string currentComputerUser = Environment.UserName.Trim();
            
            foreach (DataRow row in dtUsers.Rows)
            {
                string allowedUser = row["使用者帳號"]?.ToString().Trim();
                if (string.Equals(currentComputerUser, allowedUser, StringComparison.OrdinalIgnoreCase))
                {
                    return true; // 驗證通過
                }
            }

            // 驗證失敗
            return false;
        }
    }
}
