/// FILE: Safety_System/PdfHelper.cs ///
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public static class PdfHelper
    {
        public static void ExportDataGridViewToPdf(DataGridView dgv, string reportTitle, string fileNamePrefix)
        {
            if (dgv.Rows.Count <= 1 && dgv.AllowUserToAddRows) 
            { 
                MessageBox.Show("目前沒有資料可供導出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                return; 
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = $"{fileNamePrefix}_{DateTime.Now:yyyyMMdd}" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Form activeForm = Form.ActiveForm;
                    if (activeForm != null) activeForm.Cursor = Cursors.WaitCursor;

                    PrintDocument pd = new PrintDocument();
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;
                    pd.DefaultPageSettings.Landscape = true;
                    pd.DefaultPageSettings.Margins = new Margins(30, 30, 40, 40);

                    int rowIndex = 0;
                    int pageNumber = 1;
                    int rowsPerPageEstimate = 20;
                    int totalPages = (int)Math.Ceiling((double)(dgv.Rows.Count) / rowsPerPageEstimate);

                    pd.PrintPage += (s, ev) =>
                    {
                        Graphics g = ev.Graphics;
                        float x = ev.MarginBounds.Left;
                        float y = ev.MarginBounds.Top;
                        float pageWidth = ev.MarginBounds.Width;

                        Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                        Font fSubTitle = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
                        Font fBody = new Font("Microsoft JhengHei UI", 9F);
                        Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

                        StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                        g.DrawString("台灣玻璃工業股份有限公司-彰濱廠", fTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 40), sfCenter);
                        y += 35;
                        g.DrawString(reportTitle, fSubTitle, Brushes.Black, new RectangleF(x, y, pageWidth, 30), sfCenter);
                        y += 30;

                        g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}", fBody, Brushes.Gray, new RectangleF(x, y, pageWidth, 25), sfLeft);
                        y += 25;

                        var visCols = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList();
                        if (visCols.Count == 0) return;

                        float totalGridWidth = visCols.Sum(c => c.Width);
                        float[] actualColWidths = new float[visCols.Count];
                        for (int i = 0; i < visCols.Count; i++)
                        {
                            actualColWidths[i] = (visCols[i].Width / totalGridWidth) * pageWidth;
                        }

                        float currX = x;
                        float rowH = 32;

                        // 畫表頭
                        for (int i = 0; i < visCols.Count; i++)
                        {
                            RectangleF rect = new RectangleF(currX, y, actualColWidths[i], rowH);
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, Rectangle.Round(rect));
                            g.DrawString(visCols[i].HeaderText, fHead, Brushes.Black, rect, sfCenter);
                            currX += actualColWidths[i];
                        }
                        y += rowH;

                        // 畫資料
                        while (rowIndex < dgv.Rows.Count)
                        {
                            if (dgv.Rows[rowIndex].IsNewRow) { rowIndex++; continue; }

                            float maxRowH = rowH;
                            for (int i = 0; i < visCols.Count; i++)
                            {
                                string val = dgv[visCols[i].Index, rowIndex].Value?.ToString() ?? "";
                                SizeF sSize = g.MeasureString(val, fBody, (int)actualColWidths[i], sfLeft);
                                if (sSize.Height + 10 > maxRowH) maxRowH = sSize.Height + 10;
                            }

                            if (y + maxRowH > ev.MarginBounds.Bottom - 30)
                            {
                                g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fBody, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);
                                pageNumber++;
                                ev.HasMorePages = true;
                                return;
                            }

                            currX = x;
                            for (int i = 0; i < visCols.Count; i++)
                            {
                                RectangleF rect = new RectangleF(currX, y, actualColWidths[i], maxRowH);
                                g.DrawRectangle(Pens.Black, Rectangle.Round(rect));
                                string val = dgv[visCols[i].Index, rowIndex].Value?.ToString() ?? "";
                                RectangleF textRect = new RectangleF(rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4);
                                g.DrawString(val, fBody, Brushes.Black, textRect, sfLeft);
                                currX += actualColWidths[i];
                            }
                            y += maxRowH;
                            rowIndex++;
                        }

                        g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fBody, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);
                        ev.HasMorePages = false;
                        rowIndex = 0;
                        pageNumber = 1;
                    };

                    try
                    {
                        pd.Print();
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                        MessageBox.Show("PDF 報表匯出完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
