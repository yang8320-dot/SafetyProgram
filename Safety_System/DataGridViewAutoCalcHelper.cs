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

        private List<ColumnFormulaDef> _customFormulas;

        public DataGridViewAutoCalcHelper(DataGridView dgv, string dbName, string tableName)
        {
            _dgv = dgv;
            _dbName = dbName;
            _tableName = tableName;
            
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
                    RecalculateTable(dt, e.RowIndex); 
                } finally {
                    _isBulkUpdating = false;
                }
            }
        }

        public void RecalculateTable(DataTable dt)
        {
            RecalculateTableInternal(dt, -1, null, null);
        }

        public void RecalculateTable(DataTable dt, int specificRowIndex)
        {
            RecalculateTableInternal(dt, specificRowIndex, null, null);
        }

        public void RecalculateTable(DataTable dt, IProgress<int> progressInt, IProgress<string> progressStr)
        {
            RecalculateTableInternal(dt, -1, progressInt, progressStr);
        }

        // ============================================================================
        // 🚀 高效能批次計算核心 
        // ============================================================================
        private void RecalculateTableInternal(DataTable dt, int specificRowIndex, IProgress<int> progressInt, IProgress<string> progressStr)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

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
            Regex priceRegex = new Regex(@"PRICE\((?<cat>[^\)]+)\)");
            Regex fieldRegex = new Regex(@"\[(.*?)\]");

            using (DataTable dtMath = new DataTable()) 
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

                    if (specificRowIndex != -1 && dt.Rows.IndexOf(row) != specificRowIndex)
                    {
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

                    // --- C. 執行自訂公式換算 ---
                    if (_customFormulas != null && _customFormulas.Count > 0)
                    {
                        foreach (var formulaDef in _customFormulas)
                        {
                            string targetCol = formulaDef.TargetCol;
                            string matchCol = formulaDef.MatchCol;
                            string sDate = formulaDef.StartDate;
                            string eDate = formulaDef.EndDate;
                            string rawFormula = formulaDef.Formula;
                            string formulaType = formulaDef.FormulaType; // 數學運算 or 文字組合

                            if (!dt.Columns.Contains(targetCol)) continue;

                            if (!string.IsNullOrEmpty(matchCol) && dt.Columns.Contains(matchCol))
                            {
                                string rowDateStr = row[matchCol]?.ToString().Trim() ?? "";
                                if (string.IsNullOrEmpty(rowDateStr)) continue; 
                                if (string.Compare(rowDateStr, sDate) < 0 || string.Compare(rowDateStr, eDate) > 0) continue;
                            }

                            string evalFormula = rawFormula;
                            bool canCompute = true;

                            // 🟢 分支：文字組合 (單純變數替換，不經過 Compute)
                            if (formulaType == "文字組合")
                            {
                                var fieldMatches = fieldRegex.Matches(evalFormula);
                                foreach (Match m in fieldMatches)
                                {
                                    string colName = m.Groups[1].Value;
                                    if (dt.Columns.Contains(colName))
                                    {
                                        string val = row[colName]?.ToString().Trim() ?? "";
                                        evalFormula = evalFormula.Replace($"[{colName}]", val);
                                    }
                                }
                                row[targetCol] = evalFormula;
                            }
                            else 
                            {
                                // 🟢 分支：數學運算 (原邏輯)
                                if (evalFormula.Contains("PRICE("))
                                {
                                    DateTime targetDate = DateTime.Today; 
                                    
                                    if (!string.IsNullOrEmpty(matchCol) && dt.Columns.Contains(matchCol)) {
                                        string dateStr = row[matchCol]?.ToString().Trim() ?? "";
                                        if (dateStr.Length == 7 && dateStr.Contains("-")) {
                                            DateTime.TryParse(dateStr + "-01", out targetDate);
                                        } else {
                                            DateTime.TryParse(dateStr, out targetDate);
                                        }
                                    }

                                    var priceMatches = priceRegex.Matches(evalFormula);
                                    foreach (Match m in priceMatches) {
                                        string category = m.Groups["cat"].Value.Trim();
                                        double unitPrice = DataManager.GetUnitPrice(category, targetDate);
                                        evalFormula = evalFormula.Replace(m.Value, unitPrice.ToString());
                                    }
                                }

                                var fieldMatches = fieldRegex.Matches(evalFormula);
                                foreach (Match m in fieldMatches)
                                {
                                    string colName = m.Groups[1].Value;
                                    if (dt.Columns.Contains(colName))
                                    {
                                        string val = row[colName]?.ToString().Replace(",", "").Trim();
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
                                            double dRes = Convert.ToDouble(result);
                                            row[targetCol] = Math.Round(dRes, 4).ToString("0.####");
                                        }
                                    } 
                                    catch {
                                        if (row[targetCol]?.ToString() != "") {
                                            row[targetCol] = ""; 
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void LockTargetColumns()
        {
            var autoLockCols = new HashSet<string>();
            if (_customFormulas != null) {
                foreach (var def in _customFormulas) {
                    autoLockCols.Add(def.TargetCol);
                }
            }

            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                if (col.Name == "星期" || col.Name.EndsWith("日統計") || col.Name.EndsWith("月統計") || col.Name.EndsWith("年統計") || autoLockCols.Contains(col.Name))
                {
                    col.ReadOnly = true;
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke;
                    col.DefaultCellStyle.ForeColor = System.Drawing.Color.DimGray;
                }
            }
        }
    }
}
