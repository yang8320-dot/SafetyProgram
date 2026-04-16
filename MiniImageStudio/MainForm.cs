/* * 功能：主介面設計 (頁籤採用一般字體，縮減上方版面)
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class MainForm : Form {
        private TabControl mainTabControl;
        
        private App_Screenshot screenshotControl;
        private App_Drawing drawingControl;
        private App_Collage collageControl;
        private App_Compress compressControl;

        // 共用標準字型
        public static Font UI_Font = new Font("Microsoft JhengHei UI", 10.5f, FontStyle.Regular);

        public MainForm() {
            InitializeComponent();
        }

        private void InitializeComponent() {
            this.Text = "圖片工具程式";
            this.Size = new Size(1250, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1050, 750);
            this.BackColor = SystemColors.Control;

            // 建立主要的 TabControl
            mainTabControl = new TabControl {
                Dock = DockStyle.Fill,
                Font = UI_Font, // 【修正】採用一般字體
                ItemSize = new Size(100, 32), // 【修正】縮減頁籤的寬度與高度
                SizeMode = TabSizeMode.Fixed, 
                Padding = new Point(10, 5)
            };

            // 建立各個頁籤分頁 (TabPage)
            TabPage tabScreenshot = new TabPage("截圖") { BackColor = Color.White };
            TabPage tabDrawing = new TabPage("繪製") { BackColor = Color.White };
            TabPage tabCollage = new TabPage("拼貼") { BackColor = Color.White };
            TabPage tabCompress = new TabPage("壓縮") { BackColor = Color.White };

            // 初始化各個功能模組，並設定 Dock 填滿，然後加入對應的頁籤中
            screenshotControl = new App_Screenshot(this) { Dock = DockStyle.Fill };
            tabScreenshot.Controls.Add(screenshotControl);

            drawingControl = new App_Drawing() { Dock = DockStyle.Fill };
            tabDrawing.Controls.Add(drawingControl);

            collageControl = new App_Collage() { Dock = DockStyle.Fill };
            tabCollage.Controls.Add(collageControl);

            compressControl = new App_Compress() { Dock = DockStyle.Fill };
            tabCompress.Controls.Add(compressControl);

            // 將所有分頁加入到 TabControl 裡面
            mainTabControl.TabPages.Add(tabScreenshot);
            mainTabControl.TabPages.Add(tabDrawing);
            mainTabControl.TabPages.Add(tabCollage);
            mainTabControl.TabPages.Add(tabCompress);

            // 將 TabControl 加入到主視窗
            this.Controls.Add(mainTabControl);
        }
    }
}
