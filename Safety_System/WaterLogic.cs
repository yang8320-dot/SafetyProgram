using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class WaterLogic : DefaultLogic
    {
        public override void InitializeSchema(string dbName, string tableName)
        {
            // 水資源基礎表結構由 App_CoreTable 根據 SchemaManager 建立，此處不需額外操作
        }

        // 🟢 加入 progressInt 參數，並回報運算進度，同時優化運算效能防止卡頓
        public override async Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable dt, IProgress<int> progressInt = null, IProgress<string> progressStr = null)
        {
            await Task.Run(() =>
            {
                if (dt == null || dt.Rows.Count == 0) return;

                progressStr?.Report("正在擷取水資源歷史數據基期...");
                progressInt?.Report(5); // 🟢 回報進度
                
                string dateCol = dt.Columns.Contains("日期") ? "日期" : (dt.Columns.Contains("年月") ? "年月" : "");
                if (string.IsNullOrEmpty(dateCol)) return;

                var validRows = dt.Rows.Cast<DataRow>().Where(r => r.RowState != DataRowState.Deleted).ToList();
                if (validRows.Count == 0) return;

                string minDate = validRows.Min(r => r[dateCol]?.ToString());
                if (string.IsNullOrEmpty(minDate)) return;

                DataTable baseHistoryDt = dt.Clone(); 
                try {
                    DataTable allDbData = DataManager.GetTableData(dbName, tableName, "", "", "");
                    var baseRows = allDbData.Rows.Cast<DataRow>()
                                    .Where(r => string.Compare(r[dateCol]?.ToString(), minDate) < 0)
                                    .OrderByDescending(r => r[dateCol]?.ToString())
                                    .Take(10);
                    
                    foreach (var r in baseRows) {
                        baseHistoryDt.ImportRow(r);
                    }
                } catch { }

                progressStr?.Report("正在進行水資源差值與複合統計運算...");
                progressInt?.Report(10); // 🟢 回報進度

                DataTable calcPool = dt.Clone();
                if (baseHistoryDt.Rows.Count > 0) calcPool.Merge(baseHistoryDt);
                calcPool.Merge(dt);

                calcPool.DefaultView.Sort = $"{dateCol} ASC";
                DataTable sortedPool = calcPool.DefaultView.ToTable();

                string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };
                
                int totalSorted = sortedPool.Rows.Count;
                for (int i = 0; i < totalSorted; i++)
                {
                    // 🟢 加入進度條更新
                    if (progressInt != null && progressStr != null && (i % 20 == 0 || i == totalSorted - 1))
                    {
                        int percent = 10 + (int)((double)(i + 1) / totalSorted * 40); // 10% ~ 50%
                        progressInt.Report(percent);
                        progressStr.Report($"正在進行水資源差值運算： 第 {i + 1} 筆 / 共 {totalSorted} 筆");
                    }

                    DataRow row = sortedPool.Rows[i];
                    DataRow nextRow = (i + 1 < totalSorted) ? sortedPool.Rows[i + 1] : null;

                    foreach (DataColumn col in sortedPool.Columns)
                    {
                        string colName = col.ColumnName;

                        if (colName == "星期" && dateCol == "日期") {
                            if (DateTime.TryParse(row[dateCol]?.ToString(), out DateTime d))
                                row["星期"] = weekDays[(int)d.DayOfWeek];
                        }
                        else if (colName.EndsWith("日統計") || colName.EndsWith("月統計")) {
                            string baseCol = colName.Substring(0, colName.Length - 3);
                            if (sortedPool.Columns.Contains(baseCol)) {
                                string targetStr = row[baseCol]?.ToString().Replace(",", "").Trim();
                                string nextStr = nextRow?[baseCol]?.ToString().Replace(",", "").Trim();
                                
                                if (double.TryParse(targetStr, out double targetVal)) {
                                    if (nextRow != null && double.TryParse(nextStr, out double nextVal)) {
                                        row[colName] = (nextVal - targetVal).ToString("0.####");
                                    } else {
                                        row[colName] = "";
                                    }
                                }
                            }
                        }
                    }
                }

                progressStr?.Report("正在將運算結果映射至儲存佇列...");
                progressInt?.Report(50);
                
                int totalValid = validRows.Count;
                
                // 🟢 優化：使用 Dictionary 進行 O(1) 查找，解決大數據量時 FirstOrDefault 造成的卡頓
                var calcDict = new Dictionary<string, DataRow>();
                foreach (DataRow r in sortedPool.Rows) {
                    string d = r[dateCol]?.ToString();
                    if (!string.IsNullOrEmpty(d) && !calcDict.ContainsKey(d)) {
                        calcDict[d] = r;
                    }
                }

                for (int i = 0; i < totalValid; i++)
                {
                    // 🟢 加入進度條更新
                    if (progressInt != null && progressStr != null && (i % 20 == 0 || i == totalValid - 1))
                    {
                        int percent = 50 + (int)((double)(i + 1) / totalValid * 45); // 50% ~ 95%
                        progressInt.Report(percent);
                        progressStr.Report($"正在映射運算結果： 第 {i + 1} 筆 / 共 {totalValid} 筆");
                    }

                    DataRow targetRow = validRows[i];
                    string targetDate = targetRow[dateCol]?.ToString();
                    
                    if (targetDate != null && calcDict.TryGetValue(targetDate, out DataRow calcRow))
                    {
                        foreach (DataColumn col in dt.Columns) {
                            if (col.ColumnName.EndsWith("統計") || col.ColumnName == "星期") {
                                targetRow[col.ColumnName] = calcRow[col.ColumnName];
                            }
                        }
                    }
                }
                progressInt?.Report(95); // 保留 5% 給最終寫入
            });

            return true;
        }
    }
}
