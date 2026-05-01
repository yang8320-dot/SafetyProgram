/// FILE: Safety_System/DataGridViewAutoCalcHelper.cs ///
using System;
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

        // 🟢 UI 單筆手動輸入時的觸發邏輯
        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isBulkUpdating || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            DataGridViewRow currentRow = _dgv.Rows[e.RowIndex];
            DataGridViewRow nextRow = (e.RowIndex + 1 < _dgv.Rows.Count) ? _dgv.Rows[e.RowIndex + 1] : null;
            DataGridViewRow prevRow = (e.RowIndex - 1 >= 0) ? _dgv.Rows[e.RowIndex - 1] : null;

            // 1. 本列的日統計 = (下一列數據 - 本列數據)
            CalculateRowUI(currentRow, nextRow);

            // 2. 如果有上一列，因為本列數值變更了，所以要重新計算上一列的日統計
            if (prevRow != null) {
                CalculateRowUI(prevRow, currentRow);
            }
        }

        // 🟢 單一 UI 列的即時計算 (目標列 vs 下一列)
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
                        // 確認下一列存在且不是空白新列
                        string nextStr = (nextRow != null && !nextRow.IsNewRow) ? nextRow.Cells[baseCol].Value?.ToString().Replace(",", "").Trim() : null;
                        
                        if (double.TryParse(targetStr, out double targetVal)) {
                            // 🟢 新邏輯：下一列數值 - 本列數值
                            if (nextStr != null && double.TryParse(nextStr, out double nextVal)) {
                                targetRow.Cells[colName].Value = (nextVal - targetVal).ToString("0.####");
                            } else {
                                targetRow.Cells[colName].Value = ""; // 下一天沒資料，則當天消耗量未知
                            }
                        } else {
                            targetRow.Cells[colName].Value = "";
                        }
                    }
                }
            }
        }

        // 🚀 高效能批次計算核心：在 DataTable 記憶體中運算 (CSV匯入 / 貼上時使用)
        // 🚀 高效能批次計算核心 (支援進度條回報)
        public void RecalculateTable(DataTable dt, IProgress<int> progressInt = null, IProgress<string> progressStr = null)
        {
            if (dt == null) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                // 🟢 觸發進度條更新 (每 20 筆更新一次畫面，避免過度耗能)
                if (progressInt != null && progressStr != null && (i % 20 == 0 || i == dt.Rows.Count - 1))
                {
                    int percent = (int)((double)(i + 1) / dt.Rows.Count * 100);
                    progressInt.Report(percent);
                    progressStr.Report($"正在重新運算關聯數據： 第 {i + 1} 筆 / 共 {dt.Rows.Count} 筆 ...");
                }

                DataRow row = dt.Rows[i];
                if (row.RowState == DataRowState.Deleted) continue;
                
                // 往下尋找第一筆有效資料做為減數 (下一天)
                DataRow nextRow = null;
                for (int j = i + 1; j < dt.Rows.Count; j++) {
                    if (dt.Rows[j].RowState != DataRowState.Deleted) {
                        nextRow = dt.Rows[j];
                        break;
                    }
                }

                foreach (DataColumn col in dt.Columns)
                {
                    string colName = col.ColumnName;

                    if (colName == "星期" && dt.Columns.Contains("日期"))
                    {
                        if (DateTime.TryParse(row["日期"]?.ToString(), out DateTime d))
                            row["星期"] = weekDays[(int)d.DayOfWeek];
                        else
                            row["星期"] = "";
                    }
                    else if (colName.EndsWith("日統計") || colName.EndsWith("月統計") || colName.EndsWith("年統計"))
                    {
                        string baseCol = colName.Substring(0, colName.Length - 3);
                        if (dt.Columns.Contains(baseCol))
                        {
                            string targetStr = row[baseCol]?.ToString().Replace(",", "").Trim();
                            string nextStr = nextRow?[baseCol]?.ToString().Replace(",", "").Trim();
                            
                            if (double.TryParse(targetStr, out double targetVal)) {
                                if (nextRow != null && double.TryParse(nextStr, out double nextVal)) {
                                    row[colName] = (nextVal - targetVal).ToString("0.####");
                                } else {
                                    row[colName] = "";
                                }
                            } else {
                                row[colName] = "";
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
