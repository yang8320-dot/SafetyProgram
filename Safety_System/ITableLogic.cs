/// FILE: Safety_System/ITableLogic.cs ///
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public interface ITableLogic
    {
        void InitializeSchema(string dbName, string tableName);
        Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable savingData);
        Task OnAfterSaveAsync(string dbName, string tableName, DataTable savedData);
        string[] GetDropdownList(string tableName, string columnName);
        string[] GetDependentDropdownList(string tableName, string columnName, string parentValue);
        void OnCellClick(DataGridView dgv, DataGridViewCellEventArgs e);
    }
}
