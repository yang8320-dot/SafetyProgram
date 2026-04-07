/// FILE: Safety_System/DataGridViewAutoCalcHelper.cs ///
using System;
using System.Windows.Forms;

namespace Safety_System
{
    public class DataGridViewAutoCalcHelper
    {
        // 開關事件：用於防止大量匯入 (CSV / Ctrl+V) 時頻繁觸發 CellValueChanged 造成死機
        private bool _isBulkUpdating = false;
        private DataGridView _dgv;

        public DataGridViewAutoCalcHelper(DataGridView dgv)
        {
            _dgv = dgv;
            AttachEvents();
        }

        private void AttachEvents()
        {
            // 綁定事件前先解綁，防止重複觸發
            _dgv.DataBindingComplete -= Dgv_DataBindingComplete;
            _dgv.DataBindingComplete += Dgv_DataBindingComplete;

            _dgv.CellValueChanged -= Dgv_CellValueChanged;
            _dgv.CellValueChanged += Dgv_CellValueChanged;
        }

        // ==========================================
        // 匯入迴圈控制開關 (Begin / End)
        // ==========================================
        public void BeginBulkUpdate()
        {
            _isBulkUpdating = true;
        }

        public void EndBulkUpdate()
        {
            _isBulkUpdating = false;
            RecalculateAll(); // 匯入結束後，一次性重新計算全部
        }

        // 當資料表載入完成時，鎖定指定欄位並重新計算
        private void Dgv_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            LockTargetColumns();
            if (!_isBulkUpdating)
            {
                RecalculateAll();
            }
        }

        // 當使用者在介面上 Key in 數值時觸發
        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // 如果正在大批匯入，或點擊的是標題列，則跳過即時運算 (防呆容錯)
            if (_isBulkUpdating || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            // 觸發當前列計算
            CalculateRow(e.RowIndex);

            // 【關鍵邏輯】如果修改了第 N 列的基底數值，第 N+1 列的「統計」也會受影響，必須連鎖更新下一列
            if (e.RowIndex + 1 < _dgv.Rows.Count)
            {
                CalculateRow(e.RowIndex + 1);
            }
        }

        // ==========================================
        // 核心計算邏輯
        // ==========================================
        private void CalculateRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _dgv.Rows.Count) return;
            if (_dgv.Rows[rowIndex].IsNewRow) return;

            DataGridViewRow currentRow = _dgv.Rows[rowIndex];
            DataGridViewRow prevRow = rowIndex > 0 ? _dgv.Rows[rowIndex - 1] : null;

            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                string colName = col.Name;

                // ----------------------------------------
                // 功能 1：自動帶入星期
                // ----------------------------------------
                if (colName == "星期" && _dgv.Columns.Contains("日期"))
                {
                    string dateStr = currentRow.Cells["日期"].Value?.ToString();
                    if (DateTime.TryParse(dateStr, out DateTime dt))
                    {
                        string[] weekDays = { "日", "一", "二", "三", "四", "五", "六" };
                        currentRow.Cells["星期"].Value = weekDays[(int)dt.DayOfWeek];
                    }
                    else
                    {
                        currentRow.Cells["星期"].Value = ""; // 日期無效則清空
                    }
                }

                // ----------------------------------------
                // 功能 2：自動計算 (日統計、月統計、年統計)
                // ----------------------------------------
                if (colName.EndsWith("日統計") || colName.EndsWith("月統計") || colName.EndsWith("年統計"))
                {
                    // 取得對應的基底欄位名稱 (除去最後三個字)
                    string baseColName = colName.Substring(0, colName.Length - 3);

                    if (_dgv.Columns.Contains(baseColName))
                    {
                        // 🟢 取出目前數值，並強制濾掉千分位逗號與空白
                        string currentValStr = currentRow.Cells[baseColName].Value?.ToString().Replace(",", "").Trim();

                        // 判斷是否為有效數值 (非數值防呆容錯，非數值不運算)
                        if (double.TryParse(currentValStr, out double currentVal))
                        {
                            if (prevRow != null)
                            {
                                // 🟢 取出上一筆數值，同樣濾掉千分位逗號
                                string prevValStr = prevRow.Cells[baseColName].Value?.ToString().Replace(",", "").Trim();
                                
                                if (double.TryParse(prevValStr, out double prevVal))
                                {
                                    // 計算方法：目前值 減去 前一筆的值
                                    currentRow.Cells[colName].Value = (currentVal - prevVal).ToString("0.####");
                                }
                                else
                                {
                                    currentRow.Cells[colName].Value = ""; // 前一筆非數值無法相減
                                }
                            }
                            else
                            {
                                currentRow.Cells[colName].Value = ""; // 第一筆沒有前一筆可減，留空
                            }
                        }
                        else
                        {
                            currentRow.Cells[colName].Value = ""; // 目前值非數值
                        }
                    }
                }
            }
        }

        // ==========================================
        // 批次與初始化工具
        // ==========================================
        private void RecalculateAll()
        {
            _isBulkUpdating = true; // 暫時關閉觸發，防止迴圈
            for (int i = 0; i < _dgv.Rows.Count; i++)
            {
                CalculateRow(i);
            }
            _isBulkUpdating = false;
        }

        private void LockTargetColumns()
        {
            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                if (col.Name == "星期" || col.Name.EndsWith("日統計") || col.Name.EndsWith("月統計") || col.Name.EndsWith("年統計"))
                {
                    col.ReadOnly = true;
                    // 視覺上提示唯讀
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke; 
                    col.DefaultCellStyle.ForeColor = System.Drawing.Color.DimGray;
                }
            }
        }
    }
}
