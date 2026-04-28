/// FILE: Safety_System/WaterLogic.cs ///
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
        }

        public override async Task<bool> OnBeforeSaveAsync(string dbName, string tableName, DataTable dt)
        {
            await Task.Run(() =>
            {
                if (dt == null || dt.Rows.Count == 0) return;

                string dateCol = dt.Columns.Contains("日期") ? "日期" : (dt.Columns.Contains("年月") ? "年月" : "");
                if (string.IsNullOrEmpty(dateCol)) return;

                var validRows = dt.Rows.Cast<DataRow>().Where(r => r.RowState != DataRowState.Deleted).ToList();
                if (validRows.Count == 0) return;

                string minDate = validRows.Min(r => r[dateCol]?.ToString());
                if (string.IsNullOrEmpty(minDate)) return;

                string sql = $"SELECT * FROM [{tableName}] WHERE [{dateCol}] < '{minDate}' ORDER BY [{dateCol}] DESC LIMIT 10";
                DataTable baseHistoryDt = new DataTable();
                try {
                    DataTable allDbData = DataManager.GetTableData(dbName, tableName, "", "", "");
                    var baseRows = allDbData.Rows.Cast<DataRow>()
                                    .Where(r => string.Compare(r[dateCol]?.ToString(), minDate) < 0)
                                    .OrderByDescending(r => r[dateCol]?.ToString())
                                    .Take(10);
                    if (baseRows.Any()) baseHistoryDt = baseRows.CopyToDataTable();
                } catch { }

                DataTable calcPool = dt.Clone();
                if (baseHistoryDt.Rows.Count > 0) calcPool.Merge(baseHistoryDt);
                calcPool.Merge(dt);

                calcPool.DefaultView.Sort = $"{dateCol} ASC";
                DataTable sortedPool = calcPool.DefaultView.ToTable();

                string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };
                
                for (int i = 0; i < sortedPool.Rows.Count; i++)
                {
                    DataRow row = sortedPool.Rows[i];
                    DataRow nextRow = (i + 1 < sortedPool.Rows.Count) ? sortedPool.Rows[i + 1] : null;

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

                foreach (DataRow targetRow in validRows)
                {
                    string targetDate = targetRow[dateCol]?.ToString();
                    var calcRow = sortedPool.Rows.Cast<DataRow>().FirstOrDefault(r => r[dateCol]?.ToString() == targetDate);
                    if (calcRow != null)
                    {
                        foreach (DataColumn col in dt.Columns) {
                            if (col.ColumnName.EndsWith("統計") || col.ColumnName == "星期") {
                                targetRow[col.ColumnName] = calcRow[col.ColumnName];
                            }
                        }
                    }
                }
            });

            return true;
        }
    }
}
