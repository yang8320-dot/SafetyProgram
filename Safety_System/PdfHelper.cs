/// FILE: Safety_System/PdfHelper.cs ///
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public static class PdfHelper
    {
        // 🟢 輔助方法：搜尋快取，回傳對應的圖示串列 (支援單選與複選)
        public static List<Image> GetIconsFromCache(string tableName, string colName, string cellValue)
        {
            List<Image> icons = new List<Image>();
            if (string.IsNullOrEmpty(cellValue)) return icons;

            bool isSingleDropdown = App_DropdownManager.DropdownCache.Keys.Any(k => k.StartsWith($"{tableName}|{colName}|"));
            bool isMultiDropdown = App_DropdownManager.MultiSelectCache.ContainsKey($"{tableName}|{colName}");

            if (isSingleDropdown)
            {
                string prefix = $"{tableName}|{colName}|";
                foreach (var kvp in App_DropdownManager.DropdownCache)
                {
                    if (kvp.Key.StartsWith(prefix))
                    {
                        // 🟢 防呆：強制 Trim() 避免隱藏空白導致找不到圖示
                        var match = kvp.Value.FirstOrDefault(d => d.Text.Trim() == cellValue.Trim());
                        if (match != null && !string.IsNullOrEmpty(match.IconBase64))
                        {
                            Image img = match.GetImage();
                            if (img != null) icons.Add(img);
                            break;
                        }
                    }
                }
            }
            else if (isMultiDropdown)
            {
                string key = $"{tableName}|{colName}";
                var opts = cellValue.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
                
                if (App_DropdownManager.MultiSelectCache.ContainsKey(key))
                {
                    var defs = App_DropdownManager.MultiSelectCache[key];
                    foreach (var opt in opts)
                    {
                        // 🟢 防呆：強制 Trim() 確保精準命中
                        var match = defs.FirstOrDefault(d => d.Text.Trim() == opt.Trim());
                        if (match != null && !string.IsNullOrEmpty(match.IconBase64))
                        {
                            Image img = match.GetImage();
                            if (img != null) icons.Add(img);
                        }
                    }
                }
            }

            return icons;
        }

        // =========================================================================================
        // DataGridView 匯出 PDF (支援純圖示與多圖並排)
        // =========================================================================================
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
                    
                    // fileNamePrefix 本質上就是 TableName
                    string tableName = fileNamePrefix; 

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
                        float baseRowH = 32 + fontSizeBonus * 2;
                        float rowH = baseRowH;

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

                        // 畫資料列
                        while (rowIndex < dgv.Rows.Count)
                        {
                            if (dgv.Rows[rowIndex].IsNewRow) { rowIndex++; continue; }

                            float maxRowH = baseRowH;
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
                                
                                string colName = visCols[i].Name;
                                string val = dgv[visCols[i].Index, rowIndex].Value?.ToString() ?? "";
                                
                                // 🟢 檢查是否有圖示
                                List<Image> cellIcons = GetIconsFromCache(tableName, colName, val);
                                
                                if (cellIcons.Count > 0)
                                {
                                    // 🟢 純圖示模式：如果有圖片，我們就不畫文字了
                                    float imgSize = 18 + fontSizeBonus; // PDF 上縮放的圖示大小
                                    float imgY = rect.Y + (rect.Height - imgSize) / 2;
                                    float imgStartX = rect.X + 4;

                                    foreach (var icon in cellIcons)
                                    {
                                        g.DrawImage(icon, imgStartX, imgY, imgSize, imgSize);
                                        imgStartX += imgSize + 4; // 水平並排間距
                                    }
                                }
                                else
                                {
                                    // 🟢 沒有圖片的話，就正常畫出文字
                                    RectangleF textRect = new RectangleF(rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4);
                                    g.DrawString(val, fBody, Brushes.Black, textRect, sfLeft);
                                }

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
        // 專供儀表板 (Dashboard) 使用的獨立 PDF 導出模板 (保持原本邏輯)
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

                        int totalPages = 1;
                        float simMargin = 40f;
                        float simW = (isLandscape ? 1169f : 827f) - (simMargin * 2); 
                        float simH = (isLandscape ? 827f : 1169f);  
                        float simBottomLimit = simH - simMargin - 30f; 
                        
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
                                if (simY == simStartY) 
                                {
                                    simScale = Math.Min(simScale, (simBottomLimit - simY) / bmp.Height);
                                    simScaledHeight = bmp.Height * simScale;
                                    simY += simScaledHeight + 20f;
                                    simIndex++;
                                }
                                else
                                {
                                    totalPages++;
                                    simY = simStartY; 
                                }
                            }
                            else
                            {
                                simY += simScaledHeight + 20f;
                                simIndex++;
                            }
                        }

                        pd.PrintPage += (s, ev) =>
                        {
                            Graphics g = ev.Graphics;
                            float w = ev.MarginBounds.Width;
                            float x = ev.MarginBounds.Left;
                            float y = ev.MarginBounds.Top; 

                            Font fTitle = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold);
                            Font fSub = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
                            Font fSign = new Font("Microsoft JhengHei UI", 12F);
                            Font fDate = new Font("Microsoft JhengHei UI", 11F);

                            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center };
                            StringFormat sfLeft = new StringFormat { Alignment = StringAlignment.Near };

                            g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fTitle, Brushes.Black, new RectangleF(x, y, w, 35), sfCenter); 
                            y += 40; 

                            g.DrawString(subTitle, fSub, Brushes.Black, new RectangleF(x, y, w, 30), sfCenter); 
                            y += 40; 

                            string sign = "廠主管：______________    經/副理：______________    課/股長：______________    制表：______________";
                            g.DrawString(sign, fSign, Brushes.Black, new RectangleF(x, y, w, 25), sfCenter); 
                            y += 35; 

                            g.DrawString(dateRangeText, fDate, Brushes.DimGray, new RectangleF(x, y, w, 20), sfLeft); 
                            y += 30; 

                            float startY = y; 
                            float bottomLimit = ev.MarginBounds.Bottom - 30; 

                            while (currentBmpIndex < bitmaps.Count)
                            {
                                Bitmap bmp = bitmaps[currentBmpIndex];
                                float scale = w / bmp.Width;
                                float scaledHeight = bmp.Height * scale;

                                if (y + scaledHeight > bottomLimit)
                                {
                                    if (y == startY) 
                                    {
                                        scale = Math.Min(scale, (float)(bottomLimit - y) / bmp.Height);
                                        scaledHeight = bmp.Height * scale;
                                        g.DrawImage(bmp, x, y, bmp.Width * scale, scaledHeight);
                                        y += scaledHeight + 20;
                                        currentBmpIndex++;
                                    }
                                    else
                                    {
                                        break; 
                                    }
                                }
                                else
                                {
                                    g.DrawImage(bmp, x, y, w, scaledHeight);
                                    y += scaledHeight + 20;
                                    currentBmpIndex++;
                                }
                            }

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
                        foreach (var bmp in bitmaps) bmp.Dispose();
                        if (activeForm != null) activeForm.Cursor = Cursors.Default;
                    }
                }
            }
        }
    }
}
