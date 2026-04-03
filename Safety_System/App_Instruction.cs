using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_Instruction
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 20, 0, 0) };
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            Label t = new Label { 
                Text = "📖 系統操作說明\n\n1. 所有資料自動存儲於 SQLite 資料庫。\n2. 刪除欄位或更改名稱需要管理員密碼 (預設: tces)。\n3. 支援 Excel 複製內容後於表格內 Ctrl+V 貼上。\n4. 若版面跑掉，請調整視窗大小，系統會自動重新排版。", 
                Font = new Font("Microsoft JhengHei UI", 16F), AutoSize = true, Location = new Point(30, 20) 
            };
            p.Controls.Add(t);
            main.Controls.Add(p);
            return main;
        }
    }
}
