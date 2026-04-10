/// FILE: Safety_System/LawRtfToCsvConverter.cs ///
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Safety_System
{
    public static class LawRtfToCsvConverter
    {
        public static void Convert(string rtfPath, string csvPath)
        {
            string plainText = "";
            
            // 利用隱藏的 RichTextBox 將 RTF 檔案解析為純文字
            using (RichTextBox rtb = new RichTextBox())
            {
                try {
                    rtb.LoadFile(rtfPath); 
                } catch {
                    // 如果檔案不是標準 RTF，嘗試以純文字載入
                    rtb.Text = File.ReadAllText(rtfPath, Encoding.Default);
                }
                plainText = rtb.Text;
            }

            string[] lines = plainText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            
            string lawName = "";
            string lawDate = ""; // 儲存解析後的法規日期 (西元格式)
            
            // 層級狀態變數
            string currentArticle = "";
            int currentXiangNumber = 0;    // 項的計數器 (數字)
            string currentItemLevel3 = ""; // 款
            string currentItemLevel4 = ""; // 目
            StringBuilder currentContent = new StringBuilder();
            
            List<string[]> records = new List<string[]>();
            
            // 辨識「修正日期：民國 xxx 年 xx 月 xx 日」或「發布日期：...」
            Regex dateRegex = new Regex(@"(修正|發布)日期：?\s*民國\s*(?<year>\d+)\s*年\s*(?<month>\d+)\s*月\s*(?<day>\d+)\s*日");
            
            // 辨識「條」：包含「第10條」、「第10-1條」、「第十條之一」
            Regex articleRegex = new Regex(@"^第\s*[一二三四五六七八九十百千\d]+\s*條(?:-\d+|之\s*[一二三四五六七八九十百千\d]+)?");
            
            // 辨識「款」：「一、」「二、」等中文數字加頓號
            Regex kuanRegex = new Regex(@"^[一二三四五六七八九十百千]+、"); 
            
            // 辨識「目」：以 (一)、(二) 等左右括號包圍的中文數字
            Regex muRegex = new Regex(@"^[(（][一二三四五六七八九十百千]+[)）]");

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                
                // 🟢 需求5：刪除行首的「數字+空格」(RTF 轉文字常見的雜訊)
                line = Regex.Replace(line, @"^\d+\s+", "");
                
                if (string.IsNullOrEmpty(line)) continue;

                // 1. 抓取法規名稱與日期
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

                // 抓取法規日期並格式化為 yyyy-MM-dd (民國轉西元)
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

                // 2. 判斷是否為新的「條」
                Match mArticle = articleRegex.Match(line);
                if (mArticle.Success)
                {
                    // 遇到新的一條，先把上一條的內容存檔
                    if (!string.IsNullOrEmpty(currentArticle) && currentContent.Length > 0)
                    {
                        records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                    }
                    
                    // 🟢 需求3：格式化「條」為 3 碼
                    currentArticle = FormatTiao(mArticle.Value);
                    
                    // 換條時，重置所有的 項、款、目 計數器
                    currentXiangNumber = 1; // 預設新的一條至少有第 1 項
                    currentItemLevel3 = "";
                    currentItemLevel4 = "";
                    currentContent.Clear();
                    
                    string remainder = line.Substring(mArticle.Length).Trim();
                    if (!string.IsNullOrEmpty(remainder)) {
                        currentContent.AppendLine(remainder);
                    }
                    continue;
                }

                // 3. 判斷是否為「項」、「款」或「目」
                if (!string.IsNullOrEmpty(currentArticle))
                {
                    Match mKuan = kuanRegex.Match(line);
                    Match mMu = muRegex.Match(line);

                    if (mKuan.Success)
                    {
                        // 遇到「款」
                        if (currentContent.Length > 0) {
                            records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                            currentContent.Clear();
                        }
                        // 🟢 需求2 & 4：轉換為阿拉伯數字並補滿 2 碼
                        currentItemLevel3 = ChineseToArabic(mKuan.Value).ToString("D2"); 
                        currentItemLevel4 = ""; // 目要清空
                        currentContent.AppendLine(line.Substring(mKuan.Length).Trim());
                    }
                    else if (mMu.Success)
                    {
                        // 遇到「目」
                        if (currentContent.Length > 0) {
                            records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                            currentContent.Clear();
                        }
                        // 🟢 需求1 & 2 & 4：辨識括號目，轉換為阿拉伯數字並補滿 2 碼
                        currentItemLevel4 = ChineseToArabic(mMu.Value).ToString("D2"); 
                        currentContent.AppendLine(line.Substring(mMu.Length).Trim());
                    }
                    else
                    {
                        // 遇到一般的法條文字 (判斷是否為新的「項」)
                        // 如果目前沒有款、目，且上一段內容以「句號(。)」或「冒號(：)」結尾，則視為一個新「項」的開始
                        if (string.IsNullOrEmpty(currentItemLevel3) && string.IsNullOrEmpty(currentItemLevel4)) 
                        {
                            string currentText = currentContent.ToString().Trim();
                            
                            if (currentText.EndsWith("。") || currentText.EndsWith("：") || currentText.EndsWith(":") || currentText.EndsWith("."))
                            {
                                if (currentContent.Length > 0) {
                                    records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentText));
                                    currentContent.Clear();
                                    
                                    currentXiangNumber++; // 項計數器 + 1
                                }
                            }
                        }
                        
                        currentContent.AppendLine(line);
                    }
                }
            }

            // 儲存檔案最後結尾的內容
            if (!string.IsNullOrEmpty(currentArticle) && currentContent.Length > 0)
            {
                records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber, currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
            }

            // 寫出為標準 CSV
            StringBuilder csv = new StringBuilder();
            
            // 標題需與 App_Law_Generic 定義的資料庫欄位 100% 吻合
            string[] headers = { "日期", "類別", "法規名稱", "條", "項", "款", "目", "內容", "重點摘要", "適用性", "有提升績效機會", "有潛在不符合風險", "鑑別日期", "備註" };
            csv.AppendLine(string.Join(",", headers));

            foreach (var rec in records)
            {
                for (int i = 0; i < rec.Length; i++)
                {
                    rec[i] = EscapeCsv(rec[i]);
                }
                csv.AppendLine(string.Join(",", rec));
            }

            File.WriteAllText(csvPath, csv.ToString(), Encoding.UTF8);
        }

        // =========================================================
        // 輔助方法區
        // =========================================================

        // 建立單筆紀錄 (加入「項」的 2 碼格式化)
        private static string[] CreateRecord(string date, string lawName, string tiao, int xiang, string kuan, string mu, string content)
        {
            string[] rec = new string[14];
            for (int i = 0; i < 14; i++) rec[i] = "";
            
            rec[0] = date;    // 日期 
            rec[2] = lawName; // 法規名稱
            rec[3] = tiao;    // 條 (已在外部格式化為 3 碼)
            rec[4] = xiang.ToString("D2"); // 🟢 需求4：項格式化為 2 碼
            rec[5] = kuan;    // 款 (已在外部格式化為 2 碼)
            rec[6] = mu;      // 目 (已在外部格式化為 2 碼)
            rec[7] = content; // 內容
            
            return rec;
        }

        // 🟢 需求3：將「第十二條之一」或「第12-1條」轉換為「012-1」的 3 碼格式
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
                    return $"{mainNum:D3}-{subNum}"; // 例如：012-1
                }
                return $"{mainNum:D3}"; // 例如：012
            }
            return articleStr; // 容錯防護
        }

        // 🟢 核心轉換器：將中文數字字串(包含雜訊)轉換為阿拉伯整數
        private static int ChineseToArabic(string chnStr)
        {
            // 去除多餘符號 (適用於款、目)
            chnStr = chnStr.Replace("、", "").Replace("(", "").Replace(")", "").Replace("（", "").Replace("）", "").Trim();
            
            // 如果本身已經是阿拉伯數字，直接回傳
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
                    if (current == 0) current = 1; // 處理「十一」-> 1*10+1
                    total += current * v;
                    current = 0;
                } else {
                    current = v;
                }
            }
            total += current;
            return total;
        }

        // CSV 格式跳脫處理
        private static string EscapeCsv(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            return field;
        }
    }
}
