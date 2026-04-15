/* * 功能：主介面設計 (取消黑色背景，優化清晰度)
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class MainForm : Form {
        private Panel menuPanel;
        private Panel contentPanel;
        
        private App_Screenshot screenshotControl;
        private App_Drawing drawingControl;
        private App_Collage collageControl;
        private App_Compress compressControl;

        // 全域清晰字型設定
        public static Font UI_Font = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);

        public MainForm() {
            InitializeComponent();
        }

        private void InitializeComponent() {
            this.Text = "Mini Image Studio - Professional Portable";
            this.Size = new Size(1100, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(900, 700);
            this.BackColor = SystemColors.Control; // 改為系統預設背景色

            // 頂部選單面板 - 取消黑色背景
            menuPanel = new Panel { 
                Dock = DockStyle.Top, 
                Height = 65, 
                BackColor = SystemColors.ControlLight, // 改為淺灰色
                BorderStyle = BorderStyle.FixedSingle
            };
            
            contentPanel = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.White 
            };

            // 建立選單按鈕
            string[] navNames = { "截圖", "繪製", "拼貼", "壓縮" };
            for (int i = 0; i < navNames.Length; i++) {
                Button btn = CreateMenuButton(navNames[i], i);
                menuPanel.Controls.Add(btn);
            }

            this.Controls.Add(contentPanel);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 }); // 間距
            this.Controls.Add(menuPanel);

            // 預設載入
            ShowModule(screenshotControl = new App_Screenshot(this));
        }

        private Button CreateMenuButton(string text, int index) {
            Button btn = new Button {
                Text = text,
                Size = new Size(110, 45),
                Location = new Point(15 + (index * 125), 10),
                FlatStyle = FlatStyle.Standard, // 使用標準樣式更清晰
                Font = UI_Font,
                Cursor = Cursors.Hand
            };

            btn.Click += (s, e) => {
                switch (text) {
                    case "截圖": ShowModule(screenshotControl ?? (screenshotControl = new App_Screenshot(this))); break;
                    case "繪製": ShowModule(drawingControl ?? (drawingControl = new App_Drawing())); break;
                    case "拼貼": ShowModule(collageControl ?? (collageControl = new App_Collage())); break;
                    case "壓縮": ShowModule(compressControl ?? (compressControl = new App_Compress())); break;
                }
            };
            return btn;
        }

        private void ShowModule(UserControl module) {
            contentPanel.Controls.Clear();
            module.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(module);
        }
    }
}
