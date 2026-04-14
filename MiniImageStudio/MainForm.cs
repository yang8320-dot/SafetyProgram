/* * 功能：主介面設計與模組切換 (頂部選單)
 * 對應選單名稱：主控台
 * 對應資料庫名稱：HistoryDB (純文字檔)
 * 對應資料表名稱：主視窗配置
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class MainForm : Form {
        private Panel menuPanel;
        private Panel contentPanel;
        
        // 模組實體
        private App_Screenshot screenshotControl;
        private App_Drawing drawingControl;
        private App_Collage collageControl;
        private App_Compress compressControl;

        public MainForm() {
            InitializeComponent();
        }

        private void InitializeComponent() {
            this.Text = "Mini Image Studio - Portable";
            this.Size = new Size(1024, 768);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(800, 600);
            this.BackColor = Color.FromArgb(45, 45, 48);

            // 頂部選單面板
            menuPanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(28, 28, 28) };
            
            // 模組容器面板
            contentPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(245, 245, 247) };

            // 建立選單按鈕
            Button btnScreenshot = CreateMenuButton("截圖", 0);
            Button btnDrawing = CreateMenuButton("繪製", 1);
            Button btnCollage = CreateMenuButton("拼貼", 2);
            Button btnCompress = CreateMenuButton("壓縮", 3);

            btnScreenshot.Click += (s, e) => ShowModule(screenshotControl ?? (screenshotControl = new App_Screenshot(this)));
            btnDrawing.Click += (s, e) => ShowModule(drawingControl ?? (drawingControl = new App_Drawing()));
            btnCollage.Click += (s, e) => ShowModule(collageControl ?? (collageControl = new App_Collage()));
            btnCompress.Click += (s, e) => ShowModule(compressControl ?? (compressControl = new App_Compress()));

            menuPanel.Controls.AddRange(new Control[] { btnScreenshot, btnDrawing, btnCollage, btnCompress });

            // 組合介面：間隔 15px
            this.Controls.Add(contentPanel);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15, BackColor = Color.FromArgb(45, 45, 48) });
            this.Controls.Add(menuPanel);

            // 預設載入截圖模組
            ShowModule(screenshotControl = new App_Screenshot(this));
        }

        private Button CreateMenuButton(string text, int index) {
            Button btn = new Button {
                Text = text,
                Size = new Size(100, 40),
                Location = new Point(10 + (index * 110), 10), // 10px 邊距
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(60, 60, 60),
                Font = new Font("Microsoft JhengHei UI", 10, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void ShowModule(UserControl module) {
            contentPanel.Controls.Clear();
            module.Dock = DockStyle.Fill;
            contentPanel.Controls.Add(module);
        }
    }
}
