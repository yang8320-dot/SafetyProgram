using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace Safety_System
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 🟢 升級版：支援動態欄寬設定、下拉選單(含超過255字元防呆處理)的 Excel 匯出
        /// </summary>
        public static void ExportToExcelOrCsv(DataTable dt, string defaultFileName, Dictionary<string, float> columnWidths = null, Dictionary<string, string[]> dropdownData = null)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料可供匯出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = defaultFileName + "_" + DateTime.Now.ToString("yyyyMMdd") })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (sfd.FilterIndex == 1) // Excel
                        {
                            using (ExcelPackage p = new ExcelPackage())
                            {
                                var ws = p.Workbook.Worksheets.Add("Data");
                                ws.DefaultRowHeight = 25;
                                ws.Cells["A1"].LoadFromDataTable(dt, true);

                                // 🟢 1. 處理表頭樣式
                                using (var range = ws.Cells[1, 1, 1, dt.Columns.Count])
                                {
                                    range.Style.Font.Bold = true;
                                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                    range.Style.WrapText = true; // 允許表頭自動換行
                                }

                                // 🟢 2. 依照傳入的 UI 寬度調整 Excel 欄寬
                                if (columnWidths != null)
                                {
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    {
                                        string colName = dt.Columns[i].ColumnName;
                                        if (columnWidths.ContainsKey(colName))
                                        {
                                            // UI 的 Pixel 轉換為 Excel 的字元寬度大約是除以 7
                                            float excelWidth = columnWidths[colName] / 7f;
                                            if (excelWidth < 10) excelWidth = 10;
                                            ws.Column(i + 1).Width = excelWidth;
                                        }
                                    }
                                }
                                else
                                {
                                    ws.Cells.AutoFitColumns();
                                }

                                // 🟢 3. 處理下拉選單寫入 (解決超過 255 字元崩潰的問題)
                                if (dropdownData != null && dropdownData.Count > 0)
                                {
                                    ExcelWorksheet hiddenWs = null;
                                    int hiddenColIndex = 1;

                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    {
                                        string colName = dt.Columns[i].ColumnName;
                                        if (dropdownData.ContainsKey(colName))
                                        {
                                            string[] items = dropdownData[colName];
                                            if (items == null || items.Length == 0) continue;

                                            // 清除空白項目，避免寫入出錯
                                            var validItems = items.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                            if (validItems.Length == 0) continue;

                                            // 計算總字元長度
                                            string joinedStr = string.Join(",", validItems);
                                            var val = ws.DataValidations.AddListValidation(ws.Cells[2, i + 1, 10000, i + 1].Address);
                                            val.ShowErrorMessage = true;
                                            val.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                                            val.ErrorTitle = "輸入錯誤";
                                            val.Error = "請從下拉選單中選擇有效的選項。";

                                            // 如果字串總長度超過 255，必須採用「隱藏工作表參照」的作法
                                            if (joinedStr.Length > 255)
                                            {
                                                if (hiddenWs == null)
                                                {
                                                    hiddenWs = p.Workbook.Worksheets.Add("HiddenDropdownData");
                                                    hiddenWs.Hidden = eWorkSheetHidden.Hidden; // 隱藏此工作表
                                                }

                                                // 將選項寫入隱藏工作表的某一行
                                                for (int r = 0; r < validItems.Length; r++)
                                                {
                                                    hiddenWs.Cells[r + 1, hiddenColIndex].Value = validItems[r];
                                                }

                                                // 將公式參照設定為隱藏工作表的該欄範圍 (例如: HiddenDropdownData!$A$1:$A$10)
                                                string addressRange = $"HiddenDropdownData!${GetExcelColumnName(hiddenColIndex)}$1:${GetExcelColumnName(hiddenColIndex)}${validItems.Length}";
                                                val.Formula.ExcelFormula = addressRange;
                                                hiddenColIndex++;
                                            }
                                            else
                                            {
                                                // 小於 255 字元，直接寫入 Formula.Values
                                                foreach (var item in validItems)
                                                {
                                                    val.Formula.Values.Add(item);
                                                }
                                            }
                                        }
                                    }
                                }

                                p.SaveAs(new FileInfo(sfd.FileName));
                            }
                        }
                        else // CSV
                        {
                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            foreach (DataRow r in dt.Rows)
                            {
                                sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("資料匯出成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // 🟢 輔助工具：將數字索引轉為 Excel 的欄位字母 (1->A, 2->B, 27->AA)
        private static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        public static async Task<DataTable> ImportToDataTableAsync(string filePath, DataTable templateDt, IProgress<int> progressInt, IProgress<string> progressStr)
        {
            DataTable importedDt = templateDt.Clone();

            await Task.Run(() =>
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null || ws.Dimension == null) return;

                    int rowCount = ws.Dimension.Rows;
                    int colCount = ws.Dimension.Columns;
                    string[] headers = new string[colCount];

                    progressStr?.Report("正在解析 Excel 標題欄位...");
                    for (int c = 1; c <= colCount; c++)
                    {
                        headers[c - 1] = ws.Cells[1, c].Text.Trim();
                    }

                    for (int r = 2; r <= rowCount; r++)
                    {
                        if (r % 10 == 0 || r == rowCount)
                        {
                            int percent = (int)((double)r / rowCount * 100);
                            progressInt?.Report(percent);
                            progressStr?.Report($"正在載入資料： 第 {r} 列 / 共 {rowCount} 列 ...");
                        }

                        DataRow nr = importedDt.NewRow();
                        bool hasData = false;

                        for (int c = 1; c <= colCount; c++)
                        {
                            string cn = headers[c - 1];
                            string val = ws.Cells[r, c].Text.Trim();
                            if (importedDt.Columns.Contains(cn) && cn != "Id" && !string.IsNullOrEmpty(val))
                            {
                                nr[cn] = val;
                                hasData = true;
                            }
                        }
                        if (hasData) importedDt.Rows.Add(nr);
                    }
                }
            });

            return importedDt;
        }
    }
}
