/* * 功能：主介面設計 (修改為圖片工具程式，優化選單高度)
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

        public static Font UI_Font = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);

        public MainForm() {
            InitializeComponent();
        }

        private void InitializeComponent() {
            // 1. 修改視窗標題
            this.Text = "圖片工具程式";
            this.Size = new Size(1200, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 700);
            this.BackColor = SystemColors.Control;

            // 增加面板高度為 80，確保按鈕文字不被遮擋
            menuPanel = new Panel { 
                Dock = DockStyle.Top, 
                Height = 80, 
                BackColor = SystemColors.ControlLight,
                BorderStyle = BorderStyle.FixedSingle
            };
            
            contentPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };

            string[] navNames = { "截圖", "繪製", "拼貼", "壓縮" };
            for (int i = 0; i < navNames.Length; i++) {
                Button btn = CreateMenuButton(navNames[i], i);
                menuPanel.Controls.Add(btn);
            }

            this.Controls.Add(contentPanel);
            // 主選單與下方頁面保持舒適距離 15px
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15 }); 
            this.Controls.Add(menuPanel);

            ShowModule(screenshotControl = new App_Screenshot(this));
        }

        private Button CreateMenuButton(string text, int index) {
            Button btn = new Button {
                Text = text,
                Size = new Size(120, 45), // 加大按鈕尺寸
                Location = new Point(20 + (index * 135), 15),
                FlatStyle = FlatStyle.Standard,
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
