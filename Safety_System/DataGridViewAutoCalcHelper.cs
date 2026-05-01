using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace Safety_System
{
    public class DataGridViewAutoCalcHelper
    {
        private bool _isBulkUpdating = false;
        private DataGridView _dgv;

        public DataGridViewAutoCalcHelper(DataGridView dgv)
        {
            _dgv = dgv;
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
        public void EndBulkUpdate() { _isBulkUpdating = false; }

        private void Dgv_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            LockTargetColumns();
        }

        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isBulkUpdating || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            DataGridViewRow currentRow = _dgv.Rows[e.RowIndex];
            DataGridViewRow nextRow = (e.RowIndex + 1 < _dgv.Rows.Count) ? _dgv.Rows[e.RowIndex + 1] : null;
            DataGridViewRow prevRow = (e.RowIndex - 1 >= 0) ? _dgv.Rows[e.RowIndex - 1] : null;

            CalculateRowUI(currentRow, nextRow);
            if (prevRow != null) {
                CalculateRowUI(prevRow, currentRow);
            }
        }

        private void CalculateRowUI(DataGridViewRow targetRow, DataGridViewRow nextRow)
        {
            if (targetRow.IsNewRow) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                string colName = col.Name;
                if (colName == "星期" && _dgv.Columns.Contains("日期"))
                {
                    if (DateTime.TryParse(targetRow.Cells["日期"].Value?.ToString(), out DateTime d))
                        targetRow.Cells["星期"].Value = weekDays[(int)d.DayOfWeek];
                    else
                        targetRow.Cells["星期"].Value = "";
                }
                else if (colName.EndsWith("日統計") || colName.EndsWith("月統計") || colName.EndsWith("年統計"))
                {
                    string baseCol = colName.Substring(0, colName.Length - 3);
                    if (_dgv.Columns.Contains(baseCol))
                    {
                        string targetStr = targetRow.Cells[baseCol].Value?.ToString().Replace(",", "").Trim();
                        string nextStr = (nextRow != null && !nextRow.IsNewRow) ? nextRow.Cells[baseCol].Value?.ToString().Replace(",", "").Trim() : null;
                        
                        if (double.TryParse(targetStr, out double targetVal)) {
                            if (nextStr != null && double.TryParse(nextStr, out double nextVal)) {
                                targetRow.Cells[colName].Value = (nextVal - targetVal).ToString("0.####");
                            } else {
                                targetRow.Cells[colName].Value = ""; 
                            }
                        } else {
                            targetRow.Cells[colName].Value = "";
                        }
                    }
                }
            }
        }

        // 🚀 高效能批次計算核心：在 DataTable 記憶體中運算
        public void RecalculateTable(DataTable dt, IProgress<int> progressInt = null, IProgress<string> progressStr = null)
        {
            if (dt == null) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            // 🟢 極速優化：預先抓出所有目標計算欄位，避免在每一列迴圈中重複掃描 Columns
            var targetCols = new List<(string ColName, string BaseCol)>();
            bool hasDateCol = dt.Columns.Contains("日期");
            bool hasWeekCol = dt.Columns.Contains("星期");

            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName.EndsWith("日統計") || col.ColumnName.EndsWith("月統計") || col.ColumnName.EndsWith("年統計"))
                {
                    string baseCol = col.ColumnName.Substring(0, col.ColumnName.Length - 3);
                    if (dt.Columns.Contains(baseCol))
                    {
                        targetCols.Add((col.ColumnName, baseCol));
                    }
                }
            }

            int totalRows = dt.Rows.Count;
            for (int i = 0; i < totalRows; i++)
            {
                // 🟢 觸發進度條更新
                if (progressInt != null && progressStr != null && (i % 50 == 0 || i == totalRows - 1))
                {
                    int percent = (int)((double)(i + 1) / totalRows * 100);
                    progressInt.Report(percent);
                    progressStr.Report($"正在重新運算關聯數據： 第 {i + 1} 筆 / 共 {totalRows} 筆 ...");
                }

                DataRow row = dt.Rows[i];
                if (row.RowState == DataRowState.Deleted) continue;
                
                if (hasWeekCol && hasDateCol)
                {
                    if (DateTime.TryParse(row["日期"]?.ToString(), out DateTime d))
                        row["星期"] = weekDays[(int)d.DayOfWeek];
                    else
                        row["星期"] = "";
                }

                if (targetCols.Count == 0) continue;

                DataRow nextRow = null;
                for (int j = i + 1; j < totalRows; j++) {
                    if (dt.Rows[j].RowState != DataRowState.Deleted) {
                        nextRow = dt.Rows[j];
                        break;
                    }
                }

                foreach (var tuple in targetCols)
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
        }

        private void LockTargetColumns()
        {
            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                if (col.Name == "星期" || col.Name.EndsWith("日統計") || col.Name.EndsWith("月統計") || col.Name.EndsWith("年統計"))
                {
                    col.ReadOnly = true;
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke;
                    col.DefaultCellStyle.ForeColor = System.Drawing.Color.DimGray;
                }
            }
        }
    }
}
