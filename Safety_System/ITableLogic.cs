/// FILE: Safety_System/ITableLogic.cs ///
using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public interface ITableLogic
    {
        void InitializeSchema(string dbName, string tableName);
        // 🟢 加入進度條參數
        Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable savingData, IProgress<string> progressStr = null);
        Task OnAfterSaveAsync(string dbName, string tableName, DataTable savedData);
        string[] GetDropdownList(string tableName, string columnName);
        string[] GetDependentDropdownList(string tableName, string columnName, string parentValue);
        void OnCellClick(DataGridView dgv, DataGridViewCellEventArgs e);
    }
}
