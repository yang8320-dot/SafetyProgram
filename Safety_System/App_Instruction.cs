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
                Text = "📖 Safety System 系統核心架構與操作指南",
                Font = new Font("Microsoft JhengHei UI", 22F, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue,
                AutoSize = true,
                Location = new Point(30, 20)
            };

            RichTextBox rtb = new RichTextBox {
                Location = new Point(35, 80),
                Size = new Size(1100, 700),
                Font = new Font("Microsoft JhengHei UI", 12F),
                ForeColor = Color.FromArgb(45, 45, 45),
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            rtb.SelectionFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
            rtb.SelectionColor = Color.Navy;
            rtb.AppendText("【 🔐 系統權限與存取控制 】\n");
            rtb.AppendText("1. 一般操作密碼 (1234)：用於刪除資料列、新增紀錄等日常維護。\n");
            rtb.AppendText("2. 系統管理密碼 (11914002)：用於資料庫路徑設定、修改欄位結構、防重寫規則設定、空間清理等核心變動。\n");
            rtb.AppendText("3. 個人隱藏選單：需透過專屬密碼解鎖，支援帳密、KPI、PBC 等敏感資訊管理，密碼可於「設定 > 密碼管理」自訂。\n\n");

            rtb.SelectionFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
            rtb.SelectionColor = Color.Navy;
            rtb.AppendText("【 ⚙️ 核心資料架構 】\n");
            rtb.AppendText("• 動態資料庫引擎：系統基於 SQLite 運作，具備「動態結構轉換」能力。於介面新增/更名欄位時，資料庫會即時同步變動，具備極高擴充性。\n");
            rtb.AppendText("• 智慧防重寫 (Upsert)：匯入 Excel 或手動儲存時，系統會依據設定的「關鍵欄位」自動判斷。若資料已存在則執行「更新(Update)」，若不存在則執行「新增(Insert)」，徹底防止重複數據。\n");
            rtb.AppendText("• 自動災備機制：系統啟動時自動檢查，距上次備份超過 7 天即自動建立還原點。管理員可設定保留份數，逾期將自動循環清理空間。\n\n");

            rtb.SelectionFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
            rtb.SelectionColor = Color.Navy;
            rtb.AppendText("【 💧 水資源自動化運算引擎 】\n");
            rtb.AppendText("• 差值補算機制：輸入水錶讀數後，系統會自動尋找「次日」數據進行相減，產出「日統計」數據。若為中間插入數據，系統會自動重新串接計算鏈。\n");
            rtb.AppendText("• 匯入接軌邏輯：匯入 Excel 時，系統會自動向資料庫調閱匯入區間前、後的基期數據，確保「日/月統計量」能與現有資料完美銜接，不產生運算斷點。\n");
            rtb.AppendText("• 複合統計：內建自定義公式引擎，可跨資料表執行 SUM / AVG 等複合運算並產出同比(YoY)分析。\n\n");

            rtb.SelectionFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
            rtb.SelectionColor = Color.Navy;
            rtb.AppendText("【 ⚖️ 法規與化學品管理 】\n");
            rtb.AppendText("• 法規自動封裝：支援將「全國法規資料庫」RTF 檔案一鍵轉換為結構化 Excel，並在存檔時自動彙整「法規目錄一覽表」，追蹤適用性與鑑別進度。\n");
            rtb.AppendText("• 化學品交叉檢索：整合 12 張法規清單，支援依「名稱」或「CAS No」進行跨庫智慧比對，自動過濾無效欄位並產出法規符核度報告。\n\n");

            rtb.SelectionFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
            rtb.SelectionColor = Color.Navy;
            rtb.AppendText("【 📁 附件管理與智慧空間優化 】\n");
            rtb.AppendText("• 自動化歸檔：附件依「庫/表/年月」自動建立實體目錄結構，保持硬碟整潔。\n");
            rtb.AppendText("• 高品質圖片壓縮：上傳 JPG/PNG 圖片時，若長邊超過 1024 像素，系統將自動進行高品質比例縮放，大幅減少伺服器空間占用，同時維持稽核辨識度。\n");
            rtb.AppendText("• 孤兒檔案清理：內建掃描引擎，可找出硬碟中未被任何資料紀錄綁定的檔案並永久刪除，釋放無效空間。\n\n");

            rtb.SelectionFont = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
            rtb.SelectionColor = Color.Navy;
            rtb.AppendText("【 💡 操作技巧與 UI 強化 】\n");
            rtb.AppendText("• 欄位記憶功能：表格欄位可自由拖拉調整寬度或隱藏。系統會記憶您的習慣，下次開啟時將自動還原最佳視覺排版。\n");
            rtb.AppendText("• A4 滿版列印：導出 PDF 報表時，系統會自動運算所有可見欄位比例，並延伸至填滿 A4 橫式面寬，確保報表專業美觀。\n");
            rtb.AppendText("• 快捷鍵支援：\n");
            rtb.AppendText("  - [Ctrl + S] 全域快速儲存數據。\n");
            rtb.AppendText("  - [Ctrl + V] 支援從 Excel 複製後直接選取儲存格貼上。\n");
            rtb.AppendText("  - [Alt + Enter] 在表格儲存格內執行強制換行。\n");

            main.Controls.Add(lblTitle);
            main.Controls.Add(rtb);

            return main;
        }
    }
}
