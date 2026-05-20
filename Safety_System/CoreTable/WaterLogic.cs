using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace Safety_System
{
    public class WaterLogic : DefaultLogic
    {
        public override void InitializeSchema(string dbName, string tableName)
        {
            // 水資源基礎表結構由 App_CoreTable 根據 SchemaManager 建立，此處不需額外操作
        }

        public override async Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable dt, IProgress<int> progressInt = null, IProgress<string> progressStr = null)
        {
            // 如果傳入的 DataTable 是空的，或沒有任何修改，直接放行
            if (dt == null || dt.Rows.Count == 0) return true;

            await Task.Run(() =>
            {
                progressStr?.Report("正在擷取水資源歷史數據基期...");
                progressInt?.Report(5); 
                
                string dateCol = dt.Columns.Contains("日期") ? "日期" : (dt.Columns.Contains("年月") ? "年月" : "");
                if (string.IsNullOrEmpty(dateCol)) return;

                // 🟢 核心優化：只針對「有異動 (Added / Modified)」的資料列進行處理
                var validRows = dt.Rows.Cast<DataRow>().Where(r => r.RowState != DataRowState.Deleted).ToList();
                if (validRows.Count == 0) return;

                string minDate = validRows.Min(r => r[dateCol]?.ToString());
                if (string.IsNullOrEmpty(minDate)) return;

                DataTable baseHistoryDt = dt.Clone(); 
                try {
                    // 🟢 效能大躍進：不要全表撈取！只利用 SQL 語法撈取「小於 minDate」的最近 5 筆紀錄當作計算基期
                    string sqlFetchBase = $"SELECT * FROM [{tableName}] WHERE [{dateCol}] < '{minDate}' ORDER BY [{dateCol}] DESC LIMIT 5";
                    
                    using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={System.IO.Path.Combine(DataManager.BasePath, dbName + ".sqlite")};Version=3;"))
                    {
                        conn.Open();
                        using (var cmd = new System.Data.SQLite.SQLiteCommand(sqlFetchBase, conn))
                        using (var da = new System.Data.SQLite.SQLiteDataAdapter(cmd))
                        {
                            da.Fill(baseHistoryDt);
                        }
                    }
                } catch { }

                progressStr?.Report("正在進行水資源差值與複合統計運算...");
                progressInt?.Report(10); 

                DataTable calcPool = dt.Clone();
                if (baseHistoryDt.Rows.Count > 0) calcPool.Merge(baseHistoryDt);
                calcPool.Merge(dt);

                // 依日期排序以利計算前後差值
                calcPool.DefaultView.Sort = $"{dateCol} ASC";
                DataTable sortedPool = calcPool.DefaultView.ToTable();

                string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };
                
                int totalSorted = sortedPool.Rows.Count;
                for (int i = 0; i < totalSorted; i++)
                {
                    if (progressInt != null && progressStr != null && (i % 20 == 0 || i == totalSorted - 1))
                    {
                        int percent = 10 + (int)((double)(i + 1) / totalSorted * 40); 
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
                
                // 將計算結果暫存為 Dictionary，供快速映射回原 DataTable (O(1) 搜尋)
                var calcDict = new Dictionary<string, DataRow>();
                foreach (DataRow r in sortedPool.Rows) {
                    string d = r[dateCol]?.ToString();
                    if (!string.IsNullOrEmpty(d) && !calcDict.ContainsKey(d)) {
                        calcDict[d] = r;
                    }
                }

                for (int i = 0; i < totalValid; i++)
                {
                    if (progressInt != null && progressStr != null && (i % 20 == 0 || i == totalValid - 1))
                    {
                        int percent = 50 + (int)((double)(i + 1) / totalValid * 45); 
                        progressInt.Report(percent);
                        progressStr.Report($"正在映射運算結果： 第 {i + 1} 筆 / 共 {totalValid} 筆");
                    }

                    DataRow targetRow = validRows[i];
                    string targetDate = targetRow[dateCol]?.ToString();
                    
                    if (targetDate != null && calcDict.TryGetValue(targetDate, out DataRow calcRow))
                    {
                        foreach (DataColumn col in dt.Columns) {
                            // 只映射有被計算出來的「統計欄位」或「星期」
                            if (col.ColumnName.EndsWith("統計") || col.ColumnName == "星期") {
                                targetRow[col.ColumnName] = calcRow[col.ColumnName];
                            }
                        }
                    }
                }
                progressInt?.Report(95); 
            });

            return true;
        }
    }
}
