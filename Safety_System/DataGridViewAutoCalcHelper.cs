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

        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isBulkUpdating || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            CalculateRowUI(_dgv.Rows[e.RowIndex], e.RowIndex > 0 ? _dgv.Rows[e.RowIndex - 1] : null);
            if (e.RowIndex + 1 < _dgv.Rows.Count) {
                CalculateRowUI(_dgv.Rows[e.RowIndex + 1], _dgv.Rows[e.RowIndex]);
            }
        }

        // 針對單一 UI 列的即時計算 (User 手動 Key in 時使用)
        private void CalculateRowUI(DataGridViewRow currentRow, DataGridViewRow prevRow)
        {
            if (currentRow.IsNewRow) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                string colName = col.Name;
                if (colName == "星期" && _dgv.Columns.Contains("日期"))
                {
                    if (DateTime.TryParse(currentRow.Cells["日期"].Value?.ToString(), out DateTime d))
                        currentRow.Cells["星期"].Value = weekDays[(int)d.DayOfWeek];
                    else
                        currentRow.Cells["星期"].Value = "";
                }
                else if (colName.EndsWith("日統計") || colName.EndsWith("月統計") || colName.EndsWith("年統計"))
                {
                    string baseCol = colName.Substring(0, colName.Length - 3);
                    if (_dgv.Columns.Contains(baseCol))
                    {
                        string curStr = currentRow.Cells[baseCol].Value?.ToString().Replace(",", "").Trim();
                        string prevStr = prevRow?.Cells[baseCol].Value?.ToString().Replace(",", "").Trim();
                        
                        if (double.TryParse(curStr, out double curVal)) {
                            if (prevRow != null && double.TryParse(prevStr, out double prevVal)) {
                                currentRow.Cells[colName].Value = (curVal - prevVal).ToString("0.####");
                            } else {
                                currentRow.Cells[colName].Value = "";
                            }
                        } else {
                            currentRow.Cells[colName].Value = "";
                        }
                    }
                }
            }
        }

        // 🚀 高效能批次計算核心：直接在 DataTable 記憶體中運算 (CSV匯入 / 貼上時使用)
        public void RecalculateTable(DataTable dt)
        {
            if (dt == null) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                if (row.RowState == DataRowState.Deleted) continue;
                
                DataRow prevRow = null;
                for (int j = i - 1; j >= 0; j--) {
                    if (dt.Rows[j].RowState != DataRowState.Deleted) {
                        prevRow = dt.Rows[j];
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
                            string curStr = row[baseCol]?.ToString().Replace(",", "").Trim();
                            string prevStr = prevRow?[baseCol]?.ToString().Replace(",", "").Trim();
                            
                            if (double.TryParse(curStr, out double curVal)) {
                                if (prevRow != null && double.TryParse(prevStr, out double prevVal)) {
                                    row[colName] = (curVal - prevVal).ToString("0.####");
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
