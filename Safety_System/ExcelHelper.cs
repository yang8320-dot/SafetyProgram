/// FILE: Safety_System/ExcelHelper.cs ///
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public static class ExcelHelper
    {
        public static void ExportToExcelOrCsv(DataTable dt, string defaultFileName)
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
                                ws.DefaultColWidth = 15;
                                ws.Cells["A1"].LoadFromDataTable(dt, true);
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
                        // 回報進度條
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
