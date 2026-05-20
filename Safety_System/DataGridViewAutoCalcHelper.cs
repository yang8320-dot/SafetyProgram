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

            // 🟢 絕對時間序防護：
            // 當使用者手動修改某個水錶度數時，因為會影響到「上一次」與「這一次」的差值
            // 為防止使用者在畫面上反向排序造成誤算，直接在記憶體中快速重算一次整份表的時間軸
            DataTable dt = _dgv.DataSource as DataTable;
            if (dt != null)
            {
                _isBulkUpdating = true;
                try {
                    RecalculateTable(dt);
                } finally {
                    _isBulkUpdating = false;
                }
            }
        }

        // 🚀 高效能時間軸批次計算核心 (絕對時間序平滑計算)
        public void RecalculateTable(DataTable dt, IProgress<int> progressInt = null, IProgress<string> progressStr = null)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };

            // 1. 抓出所有目標計算欄位
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

            // 2. 🟢 獲取所有未刪除的列，並嚴格依照「日期」進行升冪排序 (完全無視 UI 畫面的排列順序)
            var validRows = new List<DataRow>();
            foreach (DataRow r in dt.Rows) 
            {
                if (r.RowState != DataRowState.Deleted) validRows.Add(r);
            }

            if (hasDateCol)
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
            
            // 3. 執行時間軸差值運算 (永遠拿時間軸上的「下一筆實際紀錄」來減)
            for (int i = 0; i < totalRows; i++)
            {
                if (progressInt != null && progressStr != null && (i % 50 == 0 || i == totalRows - 1))
                {
                    int percent = (int)((double)(i + 1) / totalRows * 100);
                    progressInt.Report(percent);
                    progressStr.Report($"正在進行時間軸關聯重算： 第 {i + 1} 筆 / 共 {totalRows} 筆 ...");
                }

                DataRow row = validRows[i];
                DataRow nextRow = (i + 1 < totalRows) ? validRows[i + 1] : null;

                if (hasWeekCol && hasDateCol)
                {
                    if (DateTime.TryParse(row["日期"]?.ToString(), out DateTime d))
                        row["星期"] = weekDays[(int)d.DayOfWeek];
                    else
                        row["星期"] = "";
                }

                if (targetCols.Count == 0) continue;

                foreach (var tuple in targetCols)
                {
                    string targetStr = row[tuple.BaseCol]?.ToString().Replace(",", "").Trim();
                    string nextStr = nextRow?[tuple.BaseCol]?.ToString().Replace(",", "").Trim();
                    
                    if (double.TryParse(targetStr, out double targetVal)) {
                        if (nextRow != null && double.TryParse(nextStr, out double nextVal)) {
                            // 直接相減，確保總消耗量正確紀錄
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
