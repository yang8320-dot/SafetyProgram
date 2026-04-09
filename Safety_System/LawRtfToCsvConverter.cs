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
                rtb.LoadFile(rtfPath); 
                plainText = rtb.Text;
            }

            string[] lines = plainText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            
            string lawName = "";
            string currentArticle = "";
            StringBuilder currentContent = new StringBuilder();
            
            List<string[]> records = new List<string[]>();
            
            // 辨識「第 OOO 條」的正規表達式 (支援 第 10-1 條 這種寫法)
            Regex articleRegex = new Regex(@"^第\s*[一二三四五六七八九十百千\d]+\s*條(-\d+)?");
            
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrEmpty(line)) continue;

                // 智能抓取法規名稱 (通常是第一行，且不以"第"開頭，長度不會太長)
                if (string.IsNullOrEmpty(lawName) && !line.StartsWith("第") && line.Length < 50)
                {
                    lawName = line;
                    continue;
                }

                Match m = articleRegex.Match(line);
                if (m.Success)
                {
                    // 遇到新的一條，先把上一條的內容存起來
                    if (!string.IsNullOrEmpty(currentArticle))
                    {
                        records.Add(CreateRecord(lawName, currentArticle, currentContent.ToString().Trim()));
                    }
                    
                    currentArticle = m.Value.Trim();
                    currentContent.Clear();
                    
                    // 抓取「第 X 條」後面的剩餘文字(可能是內容)
                    string remainder = line.Substring(m.Length).Trim();
                    if (!string.IsNullOrEmpty(remainder)) {
                        currentContent.AppendLine(remainder);
                    }
                }
                else
                {
                    // 屬於當前法條的內容 (項、款、目等)
                    if (!string.IsNullOrEmpty(currentArticle))
                    {
                        currentContent.AppendLine(line);
                    }
                }
            }

            // 儲存最後一條
            if (!string.IsNullOrEmpty(currentArticle))
            {
                records.Add(CreateRecord(lawName, currentArticle, currentContent.ToString().Trim()));
            }

            // 寫出為標準 CSV
            StringBuilder csv = new StringBuilder();
            
            // 對應 App_Law_Generic 資料庫的標準欄位
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

        private static string[] CreateRecord(string lawName, string article, string content)
        {
            string[] rec = new string[14];
            for (int i = 0; i < 14; i++) rec[i] = "";
            
            rec[2] = lawName; // 法規名稱
            rec[3] = ExtractNumber(article); // 條
            rec[7] = content; // 內容
            
            return rec;
        }

        // 從「第 12 條」中提取數字 "12"，若是純中文數字則保留原樣
        private static string ExtractNumber(string articleStr)
        {
            Match m = Regex.Match(articleStr, @"\d+(-\d+)?");
            if (m.Success) return m.Value;
            return articleStr; 
        }

        // CSV 格式跳脫處理 (遇到逗號、引號、換行時，用雙引號包覆並將內部的雙引號 doubled)
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
