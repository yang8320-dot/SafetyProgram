/// FILE: Safety_System/PdfHelper.cs ///
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public static class PdfHelper
    {
        // 原有的 DataGridView 匯出功能保留不變
        public static void ExportDataGridViewToPdf(DataGridView dgv, string reportTitle, string fileNamePrefix, bool isA3 = false, bool isLandscape = true)
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
                    
                    pd.DefaultPageSettings.Landscape = isLandscape;

                    if (isA3) pd.DefaultPageSettings.PaperSize = new PaperSize("A3", 1169, 1654);
                    else pd.DefaultPageSettings.PaperSize = new PaperSize("A4", 827, 1169);

                    pd.DefaultPageSettings.Margins = new Margins(30, 30, 40, 40);

                    int rowIndex = 0;
                    int pageNumber = 1;
                    
                    pd.PrintPage += (s, ev) =>
                    {
                        Graphics g = ev.Graphics;
                        float x = ev.MarginBounds.Left;
                        float y = ev.MarginBounds.Top;
                        float pageWidth = ev.MarginBounds.Width;

                        float fontSizeBonus = isA3 ? 2F : 0F;

                        Font fTitle = new Font("Microsoft JhengHei UI", 18F + fontSizeBonus, FontStyle.Bold);
                        Font fSubTitle = new Font("Microsoft JhengHei UI", 14F + fontSizeBonus, FontStyle.Bold);
                        Font fBody = new Font("Microsoft JhengHei UI", 9F + fontSizeBonus);
                        Font fHead = new Font("Microsoft JhengHei UI", 9F + fontSizeBonus, FontStyle.Bold);

                        StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

                        g.DrawString("台灣玻璃工業股份有限公司-彰濱廠", fTitle, Brushes.MidnightBlue, new RectangleF(x, y, pageWidth, 40), sfCenter);
                        y += 35 + fontSizeBonus;
                        g.DrawString(reportTitle, fSubTitle, Brushes.Black, new RectangleF(x, y, pageWidth, 30), sfCenter);
                        y += 30 + fontSizeBonus;

                        g.DrawString($"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}", fBody, Brushes.Gray, new RectangleF(x, y, pageWidth, 25), sfLeft);
                        y += 25 + fontSizeBonus;

                        var visCols = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).OrderBy(c => c.DisplayIndex).ToList();
                        if (visCols.Count == 0) return;

                        float totalGridWidth = visCols.Sum(c => c.Width);
                        float[] actualColWidths = new float[visCols.Count];
                        for (int i = 0; i < visCols.Count; i++)
                        {
                            actualColWidths[i] = (visCols[i].Width / totalGridWidth) * pageWidth;
                        }

                        float currX = x;
                        float rowH = 32 + fontSizeBonus * 2;

                        for (int i = 0; i < visCols.Count; i++)
                        {
                            RectangleF rect = new RectangleF(currX, y, actualColWidths[i], rowH);
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Black, Rectangle.Round(rect));
                            g.DrawString(visCols[i].HeaderText, fHead, Brushes.Black, rect, sfCenter);
                            currX += actualColWidths[i];
                        }
                        y += rowH;

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
                                g.DrawString($"- {pageNumber} -", fBody, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);
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

                        g.DrawString($"- {pageNumber} -", fBody, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom, pageWidth, 20), sfCenter);
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

        // =========================================================================================
        // 🟢 專供儀表板 (Dashboard) 使用的獨立 PDF 導出模板 (已修復總頁數計算精準度)
        // =========================================================================================
        public static void ExportDashboardToPdf(List<Bitmap> bitmaps, string subTitle, string dateRangeText, string defaultFileName, bool isLandscape = true)
        {
            if (bitmaps == null || bitmaps.Count == 0)
            {
                MessageBox.Show("無資料可供導出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = defaultFileName + "_" + DateTime.Now.ToString("yyyyMMdd") })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Form activeForm = Form.ActiveForm;
                    if (activeForm != null) activeForm.Cursor = Cursors.WaitCursor;

                    try
                    {
                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = isLandscape;
                        pd.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                        int currentBmpIndex = 0;
                        int pageNumber = 1;

                        // ====== 先計算精準的總頁數 ======
                        int totalPages = 1;
                        float simMargin = 40f;
                        float simW = (isLandscape ? 1169f : 827f) - (simMargin * 2); 
                        float simH = (isLandscape ? 827f : 1169f);  
                        float simBottomLimit = simH - simMargin - 30f; // 留給底部頁碼的空間
                        
                        // 🟢 核心修復：這必須與下方繪製的 Y 軸累加高度完全一致
                        // MarginTop(40) + y+=40 + y+=40 + y+=35 + y+=30 = 185f
                        float simStartY = simMargin + 40f + 40f + 35f + 30f; 
                        float simY = simStartY;

                        int simIndex = 0;
                        while (simIndex < bitmaps.Count)
                        {
                            Bitmap bmp = bitmaps[simIndex];
                            float simScale = simW / bmp.Width;
                            float simScaledHeight = bmp.Height * simScale;

                            if (simY + simScaledHeight > simBottomLimit)
                            {
                                if (simY == simStartY) // 如果單張圖高度就超過一頁，硬塞進本頁並強制壓縮
                                {
                                    simScale = Math.Min(simScale, (simBottomLimit - simY) / bmp.Height);
                                    simScaledHeight = bmp.Height * simScale;
                                    simY += simScaledHeight + 20f;
                                    simIndex++;
                                }
                                else
                                {
                                    // 超出空間，換頁處理
                                    totalPages++;
                                    simY = simStartY; // 重置下一頁的 Y 座標
                                    // 注意：這裡【不增加】 simIndex，讓這張圖在「下一頁」重新計算是否放得下
                                }
                            }
                            else
                            {
                                simY += simScaledHeight + 20f;
                                simIndex++;
                            }
                        }

                        // ====== 正式繪製 ======
                        pd.PrintPage += (s, ev) =>
                        {
                            Graphics g = ev.Graphics;
                            float w = ev.MarginBounds.Width;
                            float x = ev.MarginBounds.Left;
                            float y = ev.MarginBounds.Top; // 40

                            Font fTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
                            Font fSub = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                            Font fSign = new Font("Microsoft JhengHei UI", 12F);
                            Font fDate = new Font("Microsoft JhengHei UI", 11F);

                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center };
                            StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near };

                            // 第一行：大標題置中
                            g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
                            y += 40; // 80

                            // 第二行：副標題置中
                            g.DrawString(subTitle, fSub, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
                            y += 40; // 120

                            // 第三行：簽核置中
                            string sign = "廠主管：______________    經/副理：______________    課/股長：______________    制表：______________";
                            g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
                            y += 35; // 155

                            // 第四行：查詢區間靠左
                            g.DrawString(dateRangeText, fDate, Brushes.DimGray, new RectangleF(x, y, w, 20), sfLeft); 
                            y += 30; // 185

                            float startY = y; // 確保 startY == 185
                            float bottomLimit = ev.MarginBounds.Bottom - 30; 

                            while (currentBmpIndex < bitmaps.Count)
                            {
                                Bitmap bmp = bitmaps[currentBmpIndex];
                                float scale = w / bmp.Width;
                                float scaledHeight = bmp.Height * scale;

                                if (y + scaledHeight > bottomLimit)
                                {
                                    if (y == startY) // 圖片太高，整頁塞不下只好再縮小
                                    {
                                        scale = Math.Min(scale, (float)(bottomLimit - y) / bmp.Height);
                                        scaledHeight = bmp.Height * scale;
                                        g.DrawImage(bmp, x, y, bmp.Width * scale, scaledHeight);
                                        y += scaledHeight + 20;
                                        currentBmpIndex++;
                                    }
                                    else
                                    {
                                        break; // 換頁
                                    }
                                }
                                else
                                {
                                    g.DrawImage(bmp, x, y, w, scaledHeight);
                                    y += scaledHeight + 20;
                                    currentBmpIndex++;
                                }
                            }

                            // 底部：第?頁/共?頁置中
                            g.DrawString($"第 {pageNumber} 頁 / 共 {totalPages} 頁", fDate, Brushes.Black, new RectangleF(x, ev.MarginBounds.Bottom - 15, w, 20), sfCenter);

                            if (currentBmpIndex < bitmaps.Count)
                            {
                                pageNumber++;
                                ev.HasMorePages = true;
                            }
                            else
                            {
                                ev.HasMorePages = false;
                            }
                        };

                        pd.Print();
                        MessageBox.Show("PDF 匯出成功", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        // 統一在這裡進行 Bitmap 清理與滑鼠還原，乾淨俐落
                        foreach (var bmp in bitmaps) bmp.Dispose();
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                    }
                }
            }
        }
    }
}
