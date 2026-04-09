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
            
            // 辨識「第 OOO 條」 (包含 第10-1條)
            Regex articleRegex = new Regex(@"^第\s*[一二三四五六七八九十百千\d]+\s*條(-\d+)?");
            
            // 辨識 款、目 的正規表達式
            // 款：「一、」「二、」... 等中文數字加頓號
            Regex kuanRegex = new Regex(@"^[一二三四五六七八九十百千]+、"); 
            // 目：「(一)」「(二)」... 或「1.」「2.」或全形「１、」
            Regex muRegex = new Regex(@"^(\([一二三四五六七八九十百千]+\)|[0-9]+[、.]|[０-９]+[、.])");

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
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

                // 🟢 抓取法規日期並格式化為 yyyy-MM-dd (民國轉西元)
                if (string.IsNullOrEmpty(lawDate))
                {
                    Match mDate = dateRegex.Match(line);
                    if (mDate.Success)
                    {
                        if (int.TryParse(mDate.Groups["year"].Value, out int twYear))
                        {
                            int westernYear = twYear + 1911; // 民國年 + 1911 轉西元年
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
                        records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber.ToString(), currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                    }
                    
                    currentArticle = mArticle.Value.Trim();
                    
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
                            records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber.ToString(), currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                            currentContent.Clear();
                        }
                        currentItemLevel3 = mKuan.Value.Trim('、'); // 紀錄款號
                        currentItemLevel4 = ""; // 目要清空
                        currentContent.AppendLine(line.Substring(mKuan.Length).Trim());
                    }
                    else if (mMu.Success)
                    {
                        // 遇到「目」
                        if (currentContent.Length > 0) {
                            records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber.ToString(), currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
                            currentContent.Clear();
                        }
                        currentItemLevel4 = mMu.Value.Trim(); // 紀錄目號
                        currentContent.AppendLine(line.Substring(mMu.Length).Trim());
                    }
                    else
                    {
                        // 遇到一般的法條文字 (判斷是否為新的「項」)
                        
                        // 如果目前沒有款、目，且上一段內容以「句號(。)」或「冒號(：)」結尾，則視為一個新「項」的開始
                        if (string.IsNullOrEmpty(currentItemLevel3) && string.IsNullOrEmpty(currentItemLevel4)) 
                        {
                            string currentText = currentContent.ToString().Trim();
                            
                            // 判斷上一段落是否已經結束
                            if (currentText.EndsWith("。") || currentText.EndsWith("：") || currentText.EndsWith(":") || currentText.EndsWith("."))
                            {
                                // 把上一「項」的內容存起來
                                if (currentContent.Length > 0) {
                                    records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber.ToString(), currentItemLevel3, currentItemLevel4, currentText));
                                    currentContent.Clear();
                                    
                                    // 項計數器 + 1
                                    currentXiangNumber++; 
                                }
                            }
                        }
                        
                        // 將文字接續到當前的內容緩衝區
                        currentContent.AppendLine(line);
                    }
                }
            }

            // 儲存檔案最後結尾的內容
            if (!string.IsNullOrEmpty(currentArticle) && currentContent.Length > 0)
            {
                records.Add(CreateRecord(lawDate, lawName, currentArticle, currentXiangNumber.ToString(), currentItemLevel3, currentItemLevel4, currentContent.ToString().Trim()));
            }

            // 寫出為標準 CSV
            StringBuilder csv = new StringBuilder();
            
            // 對應 App_Law_Generic 資料庫的標準欄位順序
            string[] headers = { "日期", "類別", "法規名稱", "條", "項", "款", "目", "內容", "重點摘要", "適用性", "合法且有提升績效機會", "合法但潛在不符合風險", "鑑別日期", "備註" };
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

        // 封裝單筆紀錄的建立
        private static string[] CreateRecord(string date, string lawName, string article, string xiang, string kuan, string mu, string content)
        {
            string[] rec = new string[14];
            for (int i = 0; i < 14; i++) rec[i] = "";
            
            rec[0] = date;    // 🟢 日期 (西元格式: yyyy-MM-dd)
            rec[2] = lawName; // 法規名稱
            rec[3] = ExtractNumber(article); // 條
            rec[4] = xiang;   // 項 (數字 1, 2, 3...)
            rec[5] = kuan;    // 款
            rec[6] = mu;      // 目
            rec[7] = content; // 內容
            
            return rec;
        }

        // 從「第 12 條」中提取數字 "12"，若是純中文數字則保留原樣
        private static string ExtractNumber(string articleStr)
        {
            Match m = Regex.Match(articleStr, @"\d+(-\d+)?");
            if (m.Success) return m.Value;
            
            // 容錯：若無阿拉伯數字則回傳原始中文字 (去掉第與條)
            string pure = articleStr.Replace("第", "").Replace("條", "").Trim();
            return !string.IsNullOrEmpty(pure) ? pure : articleStr; 
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
