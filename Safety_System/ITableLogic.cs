using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public interface ITableLogic
    {
        void InitializeSchema(string dbName, string tableName);
        // 🟢 加入 IProgress<int> 以支援進度條回報
        Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable savingData, IProgress<int> progressInt = null, IProgress<string> progressStr = null);
        Task OnAfterSaveAsync(string dbName, string tableName, DataTable savedData);
        string[] GetDropdownList(string tableName, string columnName);
        string[] GetDependentDropdownList(string tableName, string columnName, string parentValue);
        void OnCellClick(DataGridView dgv, DataGridViewCellEventArgs e);
    }
}
