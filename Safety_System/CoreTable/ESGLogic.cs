/// FILE: Safety_System/CoreTable/ESGLogic.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class ESGLogic : DefaultLogic
    {
        public override string[] GetDropdownList(string tableName, string columnName)
        {
            // 針對 ESG指標管理 表的 ESG領域 欄位注入預設選項
            if (tableName == "ESG_Indicator" && columnName == "ESG領域")
            {
                // 優先讀取使用者是否在「下拉選單管理」中有自訂選項，若有則以使用者自訂為主
                string[] dbItems = App_DropdownManager.GetAllOptionsForColumn(tableName, columnName);
                if (dbItems != null && dbItems.Length > 1) 
                {
                    return dbItems;
                }

                // 若無自訂，回傳系統預設的 4 大類別
                return new[] { "職業安全", "健康衛生", "環境與氣侯", "消防與韌性", "" };
            }

            return base.GetDropdownList(tableName, columnName);
        }
    }
}
