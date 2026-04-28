/// FILE: Safety_System/ITableLogic.cs ///
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 核心表格的業務邏輯插件介面 (Strategy Pattern)
    /// 允許 App_CoreTable 根據不同模組動態切換計算、儲存與 UI 行為
    /// </summary>
    public interface ITableLogic
    {
        /// <summary>
        /// 介面初始化與表結構檢查 (如補齊法規的特有欄位)
        /// </summary>
        void InitializeSchema(string dbName, string tableName);

        /// <summary>
        /// 儲存前的預處理 (如：水污計算差值、法規比對重複項目改為 Update)
        /// </summary>
        Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable savingData);

        /// <summary>
        /// 儲存成功後的背景任務 (如：法規自動重算法規目錄)
        /// </summary>
        Task OnAfterSaveAsync(string dbName, string tableName, DataTable savedData);

        /// <summary>
        /// 取得該表格專屬的下拉選單內容 (取代原 SchemaManager 寫死邏輯)
        /// </summary>
        string[] GetDropdownList(string tableName, string columnName);

        /// <summary>
        /// 取得該表格專屬的連動下拉選單內容 (如：危害類型主項 -> 細項)
        /// </summary>
        string[] GetDependentDropdownList(string tableName, string columnName, string parentValue);

        /// <summary>
        /// 處理特殊欄位的點擊事件 (如：點擊「鑑別日期」自動填入今日)
        /// </summary>
        void OnCellClick(DataGridView dgv, DataGridViewCellEventArgs e);
    }
}
