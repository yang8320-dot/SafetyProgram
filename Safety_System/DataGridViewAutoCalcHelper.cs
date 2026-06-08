/// FILE: Safety_System/DataGridViewAutoCalcHelper.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Safety_System
{
    public class DataGridViewAutoCalcHelper
    {
        private bool _isBulkUpdating = false;
        private DataGridView _dgv;
        private string _dbName;
        private string _tableName;

        // 快取自訂公式，避免每次都去 DB 撈
        private Dictionary<string, string> _customFormulas;

        // 🟢 調整：建構子傳入庫名與表名，以利讀取公式
        public DataGridViewAutoCalcHelper(DataGridView dgv, string dbName, string tableName)
        {
            _dgv = dgv;
            _dbName = dbName;
            _tableName = tableName;
            
            // 載入自訂換算公式
            _customFormulas = DataManager.GetTableFormulas(_dbName, _tableName);
            
            AttachEvents();
        }

        private void AttachEvents()
        {
            _dgv.DataBindingComplete -= Dgv_DataBindingComplete;
            _dgv.DataBindingComplete += Dgv_DataBindingComplete;
            _dgv.CellValueChanged -= Dgv_CellValueChanged;
            _dgv.CellValueChanged += Dgv_CellValueChanged;
        }

        public void BeginBulkUpdate() { _isBulkUpdating = true; }
        
        public void EndBulkUpdate() 
        { 
            _isBulkUpdating = false;
            // 結束批次更新時重讀公式，確保用到最新設定
            _customFormulas = DataManager.GetTableFormulas(_dbName, _tableName);
        }

        private void Dgv_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            LockTargetColumns();
        }

        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isBulkUpdating || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            DataTable dt = _dgv.DataSource as DataTable;
            if (dt != null)
            {
                _isBulkUpdating = true;
                try {
                    // 只針對單列重算以提升效能，但時間軸差值依然需要整表重算
                    RecalculateTable(dt, e.RowIndex); 
                } finally {
                    _isBulkUpdating = false;
                }
            }
        }

        // 🚀 高效能批次計算核心 (整合：絕對時間序差值 + 🟢 自訂公式換算)
        public void RecalculateTable(DataTable dt, int specificRowIndex = -1, IProgress<int> progressInt = null, IProgress<string> progressStr = null)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            // 1. 抓出所有「日統計」類型的差值計算欄位
            var diffTargetCols = new List<(string ColName, string BaseCol)>();
            bool hasDateCol = dt.Columns.Contains("日期");
            bool hasWeekCol = dt.Columns.Contains("星期");

            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName.EndsWith("日統計") || col.ColumnName.EndsWith("月統計") || col.ColumnName.EndsWith("年統計"))
                {
                    string baseCol = col.ColumnName.Substring(0, col.ColumnName.Length - 3);
                    if (dt.Columns.Contains(baseCol))
                    {
                        diffTargetCols.Add((col.ColumnName, baseCol));
                    }
                }
            }

            var validRows = new List<DataRow>();
            foreach (DataRow r in dt.Rows) 
            {
                if (r.RowState != DataRowState.Deleted) validRows.Add(r);
            }

            // 僅在需要做時間軸差值運算時，才執行全表排序
            if (hasDateCol && diffTargetCols.Count > 0 && specificRowIndex == -1)
            {
                validRows.Sort((a, b) => {
                    DateTime da, db;
                    bool parseA = DateTime.TryParse(a["日期"]?.ToString(), out da);
                    bool parseB = DateTime.TryParse(b["日期"]?.ToString(), out db);
                    if (parseA && parseB) return da.CompareTo(db);
                    if (parseA) return -1;
                    if (parseB) return 1;
                    return 0;
                });
            }

            int totalRows = validRows.Count;
            using (DataTable dtMath = new DataTable()) // 用於執行公式
            {
                for (int i = 0; i < totalRows; i++)
                {
                    if (progressInt != null && progressStr != null && (i % 50 == 0 || i == totalRows - 1))
                    {
                        int percent = (int)((double)(i + 1) / totalRows * 100);
                        progressInt.Report(percent);
                        progressStr.Report($"正在執行欄位運算： 第 {i + 1} 筆 / 共 {totalRows} 筆 ...");
                    }

                    DataRow row = validRows[i];

                    // 如果有指定單行運算，非目標行跳過 (但若是影響到時間差值則不跳，這裡簡化處理)
                    if (specificRowIndex != -1 && dt.Rows.IndexOf(row) != specificRowIndex)
                    {
                        // 若有差值欄位，仍需全算以確保時間軸連貫
                        if (diffTargetCols.Count == 0) continue; 
                    }

                    // --- A. 基本處理：星期 ---
                    if (hasWeekCol && hasDateCol)
                    {
                        if (DateTime.TryParse(row["日期"]?.ToString(), out DateTime d))
                            row["星期"] = weekDays[(int)d.DayOfWeek];
                        else
                            row["星期"] = "";
                    }

                    // --- B. 執行時間軸差值運算 ---
                    if (diffTargetCols.Count > 0)
                    {
                        DataRow nextRow = (i + 1 < totalRows) ? validRows[i + 1] : null;
                        foreach (var tuple in diffTargetCols)
                        {
                            string targetStr = row[tuple.BaseCol]?.ToString().Replace(",", "").Trim();
                            string nextStr = nextRow?[tuple.BaseCol]?.ToString().Replace(",", "").Trim();
                            
                            if (double.TryParse(targetStr, out double targetVal)) {
                                if (nextRow != null && double.TryParse(nextStr, out double nextVal)) {
                                    row[tuple.ColName] = (nextVal - targetVal).ToString("0.####");
                                } else {
                                    row[tuple.ColName] = "";
                                }
                            } else {
                                row[tuple.ColName] = "";
                            }
                        }
                    }

                    // --- C. 🟢 執行自訂公式換算 ---
                    if (_customFormulas != null && _customFormulas.Count > 0)
                    {
                        foreach (var kvp in _customFormulas)
                        {
                            string targetCol = kvp.Key;
                            string rawFormula = kvp.Value;

                            if (!dt.Columns.Contains(targetCol)) continue;

                            string evalFormula = rawFormula;
                            bool canCompute = true;

                            // 找出公式中的 [欄位名稱] 並替換成實際數值
                            foreach (Match m in Regex.Matches(rawFormula, @"\[(.*?)\]"))
                            {
                                string colName = m.Groups[1].Value;
                                if (dt.Columns.Contains(colName))
                                {
                                    string val = row[colName]?.ToString().Replace(",", "").Trim();
                                    // 若抓不到數值，為了不讓 Compute 當機，預設當成 0 處理
                                    if (string.IsNullOrEmpty(val) || !double.TryParse(val, out _)) {
                                        val = "0"; 
                                    }
                                    evalFormula = evalFormula.Replace($"[{colName}]", val);
                                }
                                else
                                {
                                    canCompute = false; break;
                                }
                            }

                            if (canCompute)
                            {
                                try {
                                    object result = dtMath.Compute(evalFormula, null);
                                    if (result != DBNull.Value) {
                                        // 為了美觀與精確，四捨五入到小數點後 4 位
                                        double dRes = Convert.ToDouble(result);
                                        row[targetCol] = Math.Round(dRes, 4).ToString("0.####");
                                    }
                                } 
                                catch {
                                    row[targetCol] = ""; // 運算失敗(如除以零)時留空
                                }
                            }
                        }
                    }
                }
            }
        }

        private void LockTargetColumns()
        {
            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                // 自動鎖定星期、統計欄位，以及 🟢 被設定為公式輸出的目標欄位
                bool isFormulaTarget = _customFormulas != null && _customFormulas.ContainsKey(col.Name);

                if (col.Name == "星期" || col.Name.EndsWith("日統計") || col.Name.EndsWith("月統計") || col.Name.EndsWith("年統計") || isFormulaTarget)
                {
                    col.ReadOnly = true;
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke;
                    col.DefaultCellStyle.ForeColor = System.Drawing.Color.DimGray;
                }
            }
        }
    }
}
