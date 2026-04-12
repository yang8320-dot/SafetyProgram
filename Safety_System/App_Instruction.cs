/// FILE: Safety_System/App_Instruction.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_Instruction
    {
        public Control GetView()
        {
            Panel main = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.White, 
                Padding = new Padding(30) 
            };

            Label lblTitle = new Label {
                Text = "📖 Safety System 操作說明與系統總覽",
                Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue,
                AutoSize = true,
                Location = new Point(30, 20)
            };

            // 使用 RichTextBox 來容納長篇說明，支援自動換行與滾動，並允許使用者反白複製文字
            RichTextBox rtb = new RichTextBox {
                Location = new Point(35, 70),
                Size = new Size(1000, 650),
                Font = new Font("Microsoft JhengHei UI", 13F),
                ForeColor = Color.FromArgb(45, 45, 45),
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            rtb.Text = 
@"【 🔐 系統權限與密碼 】
1. 一般操作密碼：1234
   (用途：刪除單筆/多筆資料列等日常資料維護操作)
2. 系統管理員密碼：11914002
   (用途：新增/刪除資料表欄位、修改防重寫規則、變更資料庫與備份路徑等核心結構設定)

=======================================================

【 ⚙️ 系統核心機制 】
• 動態資料庫結構：介面上「新增/修改欄位」會直接修改 SQLite 資料庫結構，系統以資料庫實際存在的欄位為準，無需修改程式碼即可擴充。
• 自動災備系統：程式啟動時會自動檢查，若距上次備份超過 7 天，將自動打包備份檔。您可於「資料庫設定」中手動備份或調整保留份數 (預設保留 4 份，約一個月)。
• 智慧防重寫 (Upsert)：於設定中指定判斷欄位 (如:日期+名稱)，當大批匯入資料時，系統會自動比對並「更新(Update)」舊紀錄，防止資料庫產生重複數據。

=======================================================

【 🚀 模組功能導覽 】
• 通用操作：支援選取 Excel 內容後，在表格內點擊儲存格使用 [Ctrl + V] 快速貼上；支援匯出 Excel 以及 CSV 背景安全匯入。
• 水資源模組：內建自動運算引擎，輸入水錶讀數或大批貼上時，會自動計算「日統計 / 月統計」之差值與推算對應星期。
• 法規鑑別模組：支援將全國法規庫下載的 RTF 檔轉為 CSV；儲存時會自動彙整每部法規的「適用性」並產出年度法規目錄，支援直接匯出 PDF 報表。
• 化學品模組：提供化學品看板、快查系統、化學品要求規範及 SDS 清冊，符合職安衛與 ISO 稽核要求。
• 涵蓋領域：工安、護理、空污、水污、廢棄物、消防、檢測數據、教育訓練、法規鑑別、化學品管理，共十大領域。

=======================================================

【 💡 操作技巧與防呆 】
• 快捷儲存：全系統支援快捷鍵 [Ctrl + S] 快速觸發儲存按鈕。
• 極速運算：處理數千筆 CSV 匯入時，系統會在背景記憶體中完成運算才顯示，過程中請耐心等候，避免假當機。
• 響應式版面：若因螢幕解析度導致版面重疊，請嘗試最大化視窗，系統具備自動重新排版功能。";

            main.Controls.Add(lblTitle);
            main.Controls.Add(rtb);

            return main;
        }
    }
}
