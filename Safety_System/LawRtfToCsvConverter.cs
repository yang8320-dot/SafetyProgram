/// FILE: Safety_System/LawRtfToExcelConverter.cs ///
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation; // 引入 EPPlus 資料驗證功能

namespace Safety_System
{
    public static class LawRtfToExcelConverter
    {
        public static void Convert(string rtfPath, string excelPath)
        {
            string plainText = "";
            
            using (RichTextBox rtb = new RichTextBox())
            {
                try {
                    rtb.LoadFile(rtfPath); 
                } catch {
                    rtb.Text = File.ReadAllText(rtfPath, Encoding.Default);
                }
                plainText = rtb.Text;
            }

            string[] lines = plainText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            
            string lawName = "";
            string lawDate = ""; 
            
            string currentArticle = "";
            int currentXiangNumber = 0;    
            string currentItemLevel3 = ""; 
            string currentItemLevel4 = ""; 
            StringBuilder currentContent = new StringBuilder();
            
            List<string[]> records = new List<string[]>();
            
            Regex dateRegex = new Regex(@"(修正|發布)日期：?\s*民國\s*(?<year>\d+)\s*年\s*(?<month>\d+)\s*月\s*(?<day>\d+)\s*日");
            Regex articleRegex = new Regex(@"^第\s*[一二三四五六七八九十百千\d]+\s*條(?:-\d+|之\s*[一二三四五六七八九十百千\d]+)?");
            Regex kuanRegex = new Regex(@"^[一二三四五六七八九十百千]+、"); 
            Regex muRegex = new Regex(@"^[(（][一二三四五六七八九十百千]+[)）]");

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                line = Regex.Replace(line, @"^\d+\s+", "");
                if (string.IsNullOrEmpty(line)) continue;

                if (string.IsNullOrEmpty(lawName))
                {
                    if (line.StartsWith("法規名稱：")) {
                        lawName = line.Substring(5).Trim();
                        continue;
                    } 
                    else if (!line.StartsWith("第") && !line.Contains("日期") && line.Length < 50) {
                        lawName = line; 
                        continue;
                    }
                }

                if (string.IsNullOrEmpty(lawDate))
                {
                    Match mDate = dateRegex.Match(line);
                    if (mDate.Success)
                    {
                        if (int.TryParse(mDate.Groups["year"].Value, out int twYear))
                        {
                            int westernYear = twYear + 1911; 
                            string month = mDate.Groups["month"].Value.PadLeft(2, '0');
                            string day = mDate.Groups["day"].Value.PadLeft(2, '0');
                            lawDate = $"{westernYear}-{month}-{day}";
                            continue;
                        }
                    }
                }

                Match mArticle = articleRegex.Match(line);
                if (mArticle.Success)
                {
                    if (!string.IsNullOrEmpty(currentArticle) && currentContent.Length > 0)
                    {
                        records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                    }
                    
                    currentArticle = FormatTiao(mArticle.Value);
                    currentXiangNumber = 1; 
                    currentItemLevel3 = "";
                    currentItemLevel4 = "";
                    currentContent.Clear();
                    
                    string remainder = line.Substring(mArticle.Length).Trim();
                    if (!string.IsNullOrEmpty(remainder)) {
                        currentContent.AppendLine(remainder);
                    }
                    continue;
                }

                if (!string.IsNullOrEmpty(currentArticle))
                {
                    Match mKuan = kuanRegex.Match(line);
                    Match mMu = muRegex.Match(line);

                    if (mKuan.Success)
                    {
                        if (currentContent.Length > 0) {
                            records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                            currentContent.Clear();
                        }
                        currentItemLevel3 = ChineseToArabic(mKuan.Value).ToString("D2"); 
                        currentItemLevel4 = ""; 
                        currentContent.AppendLine(line.Substring(mKuan.Length).Trim());
                    }
                    else if (mMu.Success)
                    {
                        if (currentContent.Length > 0) {
                            records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                            currentContent.Clear();
                        }
                        currentItemLevel4 = ChineseToArabic(mMu.Value).ToString("D2"); 
                        currentContent.AppendLine(line.Substring(mMu.Length).Trim());
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(currentItemLevel3) && string.IsNullOrEmpty(currentItemLevel4)) 
                        {
                            string currentText = currentContent.ToString().Trim();
                            if (currentText.EndsWith("。") || currentText.EndsWith("：") || currentText.EndsWith(":") || currentText.EndsWith("."))
                            {
                                if (currentContent.Length > 0) {
                                    records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentText));
                                    currentContent.Clear();
                                    currentXiangNumber++; 
                                }
                            }
                        }
                        currentContent.AppendLine(line);
                    }
                }
            }

            if (!string.IsNullOrEmpty(currentArticle) && currentContent.Length > 0)
            {
                records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
            }

            // 移除了「類別」欄位，剩餘 13 個欄位
            string[] headers = { "日期", "法規名稱", "條", "項", "款", "目", "內容", "重點摘要", "適用性", "有提升績效機會", "有潛在不符合風險", "鑑別日期", "備註" };
            
            using (ExcelPackage p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("法規轉換資料");
                
                // 1. 寫入標題
                for (int c = 0; c < headers.Length; c++) {
                    ws.Cells[1, c + 1].Value = headers[c];
                    ws.Cells[1, c + 1].Style.Font.Bold = true;
                }

                // 2. 寫入資料
                for (int r = 0; r < records.Count; r++) {
                    for (int c = 0; c < records[r].Length; c++) {
                        ws.Cells[r + 2, c + 1].Value = records[r][c];
                    }
                }

                // 3. 設定自動換行與適當欄寬 (內容在第 7 欄)
                ws.Column(7).Width = 80; 
                ws.Column(7).Style.WrapText = true;
                ws.Cells.AutoFitColumns(15, 50);

                int maxRow = records.Count + 1; // 加上標題行
                
                if (maxRow >= 2) {
                    // 🟢 4. 在「適用性」欄位 (第 9 欄) 加入下拉選單防呆
                    int applicabilityColIndex = 9; 
                    var valApply = ws.DataValidations.AddListValidation(ws.Cells[2, applicabilityColIndex, maxRow, applicabilityColIndex].Address);
                    valApply.ShowErrorMessage = true;
                    valApply.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                    valApply.ErrorTitle = "輸入錯誤";
                    valApply.Error = "請從下拉選單中選擇有效的選項。";
                    valApply.Formula.Values.Add("適用");
                    valApply.Formula.Values.Add("不適用");
                    valApply.Formula.Values.Add("參考");
                    valApply.Formula.Values.Add("確認中");

                    // 🟢 5. 在「有提升績效機會」欄位 (第 10 欄) 加入下拉選單防呆
                    int perfColIndex = 10;
                    var valPerf = ws.DataValidations.AddListValidation(ws.Cells[2, perfColIndex, maxRow, perfColIndex].Address);
                    valPerf.ShowErrorMessage = true;
                    valPerf.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                    valPerf.ErrorTitle = "輸入錯誤";
                    valPerf.Error = "只能選擇空格或 v。";
                    valPerf.Formula.Values.Add(" "); // 空格選項 (Excel需要這樣設定)
                    valPerf.Formula.Values.Add("v");

                    // 🟢 6. 在「有潛在不符合風險」欄位 (第 11 欄) 加入下拉選單防呆
                    int riskColIndex = 11;
                    var valRisk = ws.DataValidations.AddListValidation(ws.Cells[2, riskColIndex, maxRow, riskColIndex].Address);
                    valRisk.ShowErrorMessage = true;
                    valRisk.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                    valRisk.ErrorTitle = "輸入錯誤";
                    valRisk.Error = "只能選擇空格或 v。";
                    valRisk.Formula.Values.Add(" "); 
                    valRisk.Formula.Values.Add("v");
                }

                p.SaveAs(new FileInfo(excelPath));
            }
        }

        // =========================================================
        // 輔助方法區
        // =========================================================
        private static string[] CreateRecord(string date, string lawName, string tiao, int xiang, string kuan, string mu, string content)
        {
            string[] rec = new string[13];
            for (int i = 0; i < 13; i++) rec[i] = "";
            
            rec[0] = date;             // 日期
            rec[1] = lawName;          // 法規名稱
            rec[2] = tiao;             // 條
            rec[3] = xiang.ToString("D2"); // 項
            rec[4] = kuan;             // 款
            rec[5] = mu;               // 目
            rec[6] = content;          // 內容
            rec[7] = "";               // 重點摘要
            rec[8] = "";               // 適用性
            rec[9] = "";               // 有提升績效機會
            rec[10] = "";              // 有潛在不符合風險
            rec[11] = DateTime.Today.ToString("yyyy-MM-dd"); // 鑑別日期
            rec[12] = "";              // 備註
            
            return rec;
        }

        private static string FormatTiao(string articleStr)
        {
            string pure = articleStr.Replace("第", "").Replace("條", "").Trim();
            string[] parts = pure.Split(new[] { '-', '之' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 0)
            {
                int mainNum = ChineseToArabic(parts[0]);
                if (parts.Length > 1)
                {
                    int subNum = ChineseToArabic(parts[1]);
                    return $"{mainNum:D3}-{subNum}"; 
                }
                return $"{mainNum:D3}"; 
            }
            return articleStr; 
        }

        private static int ChineseToArabic(string chnStr)
        {
            chnStr = chnStr.Replace("、", "").Replace("(", "").Replace(")", "").Replace("（", "").Replace("）", "").Trim();
            if (int.TryParse(chnStr, out int val)) return val;

            var dict = new Dictionary<char, int> {
                {'一', 1}, {'二', 2}, {'三', 3}, {'四', 4}, {'五', 5}, 
                {'六', 6}, {'七', 7}, {'八', 8}, {'九', 9}, {'十', 10}, 
                {'百', 100}, {'千', 1000}, {'零', 0}
            };
            
            int total = 0;
            int current = 0;
            
            for (int i = 0; i < chnStr.Length; i++) {
                if (!dict.ContainsKey(chnStr[i])) continue;
                int v = dict[chnStr[i]];
                
                if (v == 10 || v == 100 || v == 1000) {
                    if (current == 0) current = 1; 
                    total += current * v;
                    current = 0;
                } else {
                    current = v;
                }
            }
            total += current;
            return total;
        }
    }
}
