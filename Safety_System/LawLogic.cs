/// FILE: Safety_System/LawLogic.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class LawLogic : DefaultLogic
    {
        private const string DirectoryTableName = "法規目錄一覽";

        public override void InitializeSchema(string dbName, string tableName)
        {
            var existingCols = DataManager.GetColumnNames(dbName, tableName);
            if (!existingCols.Contains("有提升績效機會")) DataManager.AddColumn(dbName, tableName, "有提升績效機會");
            if (!existingCols.Contains("有潛在不符合風險")) DataManager.AddColumn(dbName, tableName, "有潛在不符合風險");
            
            // 🟢 強制補齊附件檔案，相容舊資料庫
            if (!existingCols.Contains("附件檔案")) DataManager.AddColumn(dbName, tableName, "附件檔案");

            DataManager.InitTable(dbName, DirectoryTableName, $@"CREATE TABLE IF NOT EXISTS [{DirectoryTableName}] (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, [選項類別] TEXT, [流水號] TEXT, [法規名稱] TEXT, [日期] TEXT, 
                [適用性] TEXT, [鑑別日期] TEXT, [再次確認日期] TEXT);");
        }

        public override string[] GetDropdownList(string tableName, string columnName)
        {
            if (columnName == "類別") return new[] { "法律", "命令", "行政規則", "解釋令函", "" };
            if (columnName == "適用性") return new[] { "適用", "不適用", "參考", "確認中", "" };
            if (columnName == "有提升績效機會" || columnName == "有潛在不符合風險") return new[] { "", "v" };
            
            if (columnName == "條") {
                var list = new List<string> { "" };
                for (int i = 1; i <= 500; i++) list.Add(i.ToString("D3"));
                return list.ToArray();
            }
            if (columnName == "項" || columnName == "款" || columnName == "目") {
                var list = new List<string> { "" };
                for (int i = 1; i <= 20; i++) list.Add(i.ToString("D2"));
                return list.ToArray();
            }

            return base.GetDropdownList(tableName, columnName);
        }

        public override async Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable dt)
        {
            await Task.Run(() =>
            {
                DataTable dbData = DataManager.GetTableData(dbName, tableName, "", "", "");
                var existingDict = new Dictionary<string, int>();

                foreach (DataRow dbRow in dbData.Rows) {
                    string name = dbRow["法規名稱"]?.ToString().Trim() ?? "";
                    string article = dbRow["條"]?.ToString().Trim() ?? "";
                    string content = dbRow["內容"]?.ToString().Trim() ?? "";
                    string key = $"{name}_|{article}_|{content}";
                    existingDict[key] = Convert.ToInt32(dbRow["Id"]);
                }

                foreach (DataRow row in dt.Rows) {
                    if (row.RowState == DataRowState.Deleted) continue;
                    
                    if (row.Table.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0)
                        continue;

                    string name = row["法規名稱"]?.ToString().Trim() ?? "";
                    string article = row["條"]?.ToString().Trim() ?? "";
                    string content = row["內容"]?.ToString().Trim() ?? "";
                    string key = $"{name}_|{article}_|{content}";

                    if (existingDict.ContainsKey(key)) {
                        bool isReadOnly = row.Table.Columns["Id"].ReadOnly;
                        row.Table.Columns["Id"].ReadOnly = false;
                        row["Id"] = existingDict[key];
                        row.Table.Columns["Id"].ReadOnly = isReadOnly;
                    }
                }
            });

            return true;
        }

        public override async Task OnAfterSaveAsync(string dbName, string tableName, DataTable dt)
        {
            await Task.Run(() =>
            {
                DataTable dtMain = DataManager.GetTableData(dbName, tableName, "", "", "");
                DataTable dtDirExist = DataManager.GetTableData(dbName, DirectoryTableName, "", "", "");
                var existingDirDict = new Dictionary<string, DataRow>();
                
                foreach (DataRow r in dtDirExist.Rows) {
                    if (r["選項類別"]?.ToString() == tableName) {
                        string name = r["法規名稱"]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(name)) existingDirDict[name] = r;
                    }
                }

                var grouped = new Dictionary<string, List<DataRow>>();
                foreach(DataRow r in dtMain.Rows) {
                    string name = r["法規名稱"]?.ToString().Trim();
                    if (string.IsNullOrEmpty(name)) continue;
                    if (!grouped.ContainsKey(name)) grouped[name] = new List<DataRow>();
                    grouped[name].Add(r);
                }

                DataTable dtDir = new DataTable();
                dtDir.Columns.Add("Id", typeof(int)); 
                dtDir.Columns.Add("選項類別", typeof(string));
                dtDir.Columns.Add("流水號", typeof(string));
                dtDir.Columns.Add("法規名稱", typeof(string));
                dtDir.Columns.Add("日期", typeof(string));
                dtDir.Columns.Add("適用性", typeof(string));
                dtDir.Columns.Add("鑑別日期", typeof(string));
                dtDir.Columns.Add("再次確認日期", typeof(string));

                int index = 1;
                HashSet<string> processedNames = new HashSet<string>();

                foreach(var kvp in grouped) {
                    string lawName = kvp.Key;
                    processedNames.Add(lawName);

                    string latestDate = "", latestIdenDate = "", applyStatus = "";
                    bool hasApplicable = false;

                    foreach(var row in kvp.Value) {
                        string d = row["日期"]?.ToString() ?? "";
                        string iden = row["鑑別日期"]?.ToString() ?? "";
                        string apply = row["適用性"]?.ToString() ?? "";

                        if (string.Compare(d, latestDate) > 0) latestDate = d;
                        if (string.Compare(iden, latestIdenDate) > 0) latestIdenDate = iden;
                        if (apply == "適用") hasApplicable = true;
                        if (string.IsNullOrEmpty(applyStatus)) applyStatus = apply; 
                    }

                    DataRow newRow = dtDir.NewRow();
                    newRow["選項類別"] = tableName; 
                    newRow["流水號"] = index.ToString(); 
                    newRow["法規名稱"] = lawName;
                    newRow["日期"] = latestDate;
                    newRow["適用性"] = hasApplicable ? "適用" : applyStatus; 
                    newRow["鑑別日期"] = latestIdenDate;

                    if (existingDirDict.ContainsKey(lawName)) {
                        newRow["Id"] = existingDirDict[lawName]["Id"];
                        newRow["再次確認日期"] = existingDirDict[lawName]["再次確認日期"]?.ToString();
                    }
                    dtDir.Rows.Add(newRow);
                    index++;
                }

                foreach (var kvp in existingDirDict) {
                    if (!processedNames.Contains(kvp.Key)) {
                        int idToDelete = Convert.ToInt32(kvp.Value["Id"]);
                        DataManager.DeleteRecord(dbName, DirectoryTableName, idToDelete);
                    }
                }

                DataManager.BulkSaveTable(dbName, DirectoryTableName, dtDir);
            });
        }

        public override void OnCellClick(DataGridView dgv, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && !dgv.Rows[e.RowIndex].IsNewRow) {
                if (dgv.Columns[e.ColumnIndex].Name == "鑑別日期") {
                    var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (cell.Value == DBNull.Value || string.IsNullOrWhiteSpace(cell.Value?.ToString())) {
                        cell.Value = DateTime.Today.ToString("yyyy-MM-dd");
                    }
                }
            }
        }
    }
}
