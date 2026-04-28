/// FILE: Safety_System/DefaultLogic.cs ///
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class DefaultLogic : ITableLogic
    {
        public virtual void InitializeSchema(string dbName, string tableName) { }

        public virtual Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable savingData)
        {
            return Task.FromResult(true); 
        }

        public virtual Task OnAfterSaveAsync(string dbName, string tableName, DataTable savedData)
        {
            return Task.CompletedTask; 
        }

        public virtual string[] GetDropdownList(string tableName, string columnName)
        {
            return TableSchemaManager.GetDropdownList(tableName, columnName);
        }

        public virtual string[] GetDependentDropdownList(string tableName, string columnName, string parentValue)
        {
            return TableSchemaManager.GetDependentDropdownList(tableName, columnName, parentValue);
        }

        public virtual void OnCellClick(DataGridView dgv, DataGridViewCellEventArgs e) { }
    }
}
