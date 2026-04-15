/* * 功能：主介面設計 (上方功能按鍵加上鮮艷底色)
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
            this.Text = "圖片工具程式 - 專業終極版";
            this.Size = new Size(1200, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1000, 700);
            this.BackColor = SystemColors.Control;

            menuPanel = new Panel { Dock = DockStyle.Top, Height = 80, BackColor = SystemColors.ControlLight, BorderStyle = BorderStyle.FixedSingle };
            contentPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };

            string[] navNames = { "截圖", "繪製", "拼貼", "壓縮" };
            for (int i = 0; i < navNames.Length; i++) {
                menuPanel.Controls.Add(CreateMenuButton(navNames[i], i));
            }

            this.Controls.Add(contentPanel);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15 }); 
            this.Controls.Add(menuPanel);

            ShowModule(screenshotControl = new App_Screenshot(this));
        }

        private Button CreateMenuButton(string text, int index) {
            // 鮮豔顏色陣列：紅、藍、綠、橘
            Color[] colors = { Color.FromArgb(220, 53, 69), Color.FromArgb(0, 123, 255), Color.FromArgb(40, 167, 69), Color.FromArgb(253, 126, 20) };
            
            Button btn = new Button {
                Text = text,
                Size = new Size(120, 45),
                Location = new Point(20 + (index * 135), 15),
                FlatStyle = FlatStyle.Flat,
                Font = new Font(UI_Font.FontFamily, 12, FontStyle.Bold),
                Cursor = Cursors.Hand,
                BackColor = colors[index % colors.Length],
                ForeColor = Color.White
            };
            btn.FlatAppearance.BorderSize = 0;

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
