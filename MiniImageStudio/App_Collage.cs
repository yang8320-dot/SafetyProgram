/* * 功能：進階拼貼模組 (5種樣板、自訂浮動文字框)
 * 對應選單名稱：拼貼
 * 對應資料庫名稱：HistoryDB
 * 對應資料表名稱：App_Collage
 */
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class App_Collage : UserControl {
        private PictureBox pb;
        private Image[] loadedImages = new Image[3];
        private Bitmap baseCollage; // 底圖拼貼

        // UI 控制項
        private ComboBox cbLayout;
        private TextBox txtContent;
        private TrackBar tbOpacity;
        private Button btnTextColor, btnBgColor, btnBorderColor;

        // 文字框屬性
        private string overlayText = "點擊此處拖曳文字";
        private Color textColor = Color.White;
        private Color bgColor = Color.Black;
        private Color borderColor = Color.White;
        private int bgOpacity = 128; // 0-255
        
        // 拖曳狀態
        private RectangleF textRect = new RectangleF(20, 20, 250, 60);
        private bool isDraggingText = false;
        private PointF dragOffset;

        public App_Collage() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
        }

        private void InitializeUI() {
            // --- 頂部面板：圖片與版型 ---
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = SystemColors.Control };
            
            Button btnLoad1 = new Button { Text = "載入圖1", Left = 10, Top = 10, Width = 70 };
            Button btnLoad2 = new Button { Text = "載入圖2", Left = 85, Top = 10, Width = 70 };
            Button btnLoad3 = new Button { Text = "載入圖3", Left = 160, Top = 10, Width = 70 };
            
            cbLayout = new ComboBox { Left = 240, Top = 12, Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            cbLayout.Items.AddRange(new string[] {
                "1. 上下 (2圖)", 
                "2. 左右 (2圖)", 
                "3. 上1 下2 (3圖)", 
                "4. 左1 右2 (3圖)", 
                "5. 左2 右1 (3圖)"
            });
            cbLayout.SelectedIndex = 0;

            Button btnGenerate = new Button { Text = "產生底圖", Left = 430, Top = 10, Width = 80, BackColor = Color.SteelBlue, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存拼貼", Left = 520, Top = 10, Width = 80, BackColor = Color.SeaGreen, ForeColor = Color.White };

            btnLoad1.Click += (s, e) => LoadImage(0);
            btnLoad2.Click += (s, e) => LoadImage(1);
            btnLoad3.Click += (s, e) => LoadImage(2);
            btnGenerate.Click += (s, e) => GenerateBaseCollage();
            btnSave.Click += (s, e) => SaveImage();

            topPanel.Controls.AddRange(new Control[] { btnLoad1, btnLoad2, btnLoad3, cbLayout, btnGenerate, btnSave });

            // --- 底部面板：進階文字設定 ---
            Panel bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 60, BackColor = SystemColors.ControlLight };
            
            Label lblText = new Label { Text = "文字內容:", Left = 10, Top = 20, AutoSize = true };
            txtContent = new TextBox { Left = 80, Top = 15, Width = 150, Text = overlayText };
            txtContent.TextChanged += (s, e) => { overlayText = txtContent.Text; pb.Invalidate(); };

            btnTextColor = new Button { Text = "字色", Left = 240, Top = 15, Width = 50, BackColor = textColor };
            btnBgColor = new Button { Text = "底色", Left = 295, Top = 15, Width = 50, BackColor = bgColor };
            btnBorderColor = new Button { Text = "框色", Left = 350, Top = 15, Width = 50, BackColor = borderColor };
            
            Label lblOpacity = new Label { Text = "透明度:", Left = 410, Top = 20, AutoSize = true };
            tbOpacity = new TrackBar { Left = 460, Top = 10, Width = 120, Minimum = 0, Maximum = 255, Value = bgOpacity, TickStyle = TickStyle.None };
            tbOpacity.ValueChanged += (s, e) => { bgOpacity = tbOpacity.Value; pb.Invalidate(); };

            btnTextColor.Click += (s, e) => ChooseColor(ref textColor, btnTextColor);
            btnBgColor.Click += (s, e) => ChooseColor(ref bgColor, btnBgColor);
            btnBorderColor.Click += (s, e) => ChooseColor(ref borderColor, btnBorderColor);

            bottomPanel.Controls.AddRange(new Control[] { lblText, txtContent, btnTextColor, btnBgColor, btnBorderColor, lblOpacity, tbOpacity });

            // --- 圖片預覽區 ---
            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.DarkGray, Cursor = Cursors.Hand };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;

            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Bottom, Height = 10 });
            this.Controls.Add(bottomPanel);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(topPanel);
        }

        private void ChooseColor(ref Color targetColor, Button btn) {
            using (ColorDialog cd = new ColorDialog { Color = targetColor }) {
                if (cd.ShowDialog() == DialogResult.OK) {
                    targetColor = cd.Color;
                    btn.BackColor = targetColor;
                    pb.Invalidate();
                }
            }
        }

        private void LoadImage(int index) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Images|*.jpg;*.png" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    if (loadedImages[index] != null) loadedImages[index].Dispose();
                    loadedImages[index] = Image.FromFile(ofd.FileName);
                    MessageBox.Show($"圖 {index + 1} 載入成功！");
                }
            }
        }

        private void GenerateBaseCollage() {
            // 基礎畫布大小 (可依需求擴充為動態計算，此處示範固定高解析度畫布)
            int W = 1200, H = 1200; 
            baseCollage = new Bitmap(W, H);

            using (Graphics g = Graphics.FromImage(baseCollage)) {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                int mode = cbLayout.SelectedIndex;
                Rectangle[] rects = new Rectangle[3];

                // 排版算法
                if (mode == 0) { // 上下
                    rects[0] = new Rectangle(0, 0, W, H / 2);
                    rects[1] = new Rectangle(0, H / 2, W, H / 2);
                } else if (mode == 1) { // 左右
                    rects[0] = new Rectangle(0, 0, W / 2, H);
                    rects[1] = new Rectangle(W / 2, 0, W / 2, H);
                } else if (mode == 2) { // 上1下2
                    rects[0] = new Rectangle(0, 0, W, H / 2);
                    rects[1] = new Rectangle(0, H / 2, W / 2, H / 2);
                    rects[2] = new Rectangle(W / 2, H / 2, W / 2, H / 2);
                } else if (mode == 3) { // 左1右2
                    rects[0] = new Rectangle(0, 0, W / 2, H);
                    rects[1] = new Rectangle(W / 2, 0, W / 2, H / 2);
                    rects[2] = new Rectangle(W / 2, H / 2, W / 2, H / 2);
                } else if (mode == 4) { // 左2右1
                    rects[0] = new Rectangle(0, 0, W / 2, H / 2);
                    rects[1] = new Rectangle(0, H / 2, W / 2, H / 2);
                    rects[2] = new Rectangle(W / 2, 0, W / 2, H);
                }

                // 將圖片繪製到對應區塊 (使用 Zoom 模式裁剪)
                for (int i = 0; i < 3; i++) {
                    if (loadedImages[i] != null && i < rects.Length && rects[i].Width > 0) {
                        DrawImageZoomed(g, loadedImages[i], rects[i]);
                    }
                }
            }
            pb.Invalidate(); // 觸發重繪
        }

        private void DrawImageZoomed(Graphics g, Image img, Rectangle destRect) {
            float ratio = Math.Max((float)destRect.Width / img.Width, (float)destRect.Height / img.Height);
            int newW = (int)(img.Width * ratio);
            int newH = (int)(img.Height * ratio);
            int x = destRect.X + (destRect.Width - newW) / 2;
            int y = destRect.Y + (destRect.Height - newH) / 2;
            
            g.SetClip(destRect); // 限制繪製範圍不超過框格
            g.DrawImage(img, x, y, newW, newH);
            g.ResetClip();
            g.DrawRectangle(Pens.White, destRect); // 格線
        }

        // --- 拖曳與渲染文字 ---
        private void Pb_Paint(object sender, PaintEventArgs e) {
            // 1. 畫底圖 (依比例縮放適應 PictureBox)
            if (baseCollage != null) {
                e.Graphics.DrawImage(baseCollage, GetDisplayRect());
            }

            // 2. 畫文字框 (覆蓋在上層)
            if (string.IsNullOrEmpty(overlayText)) return;
            
            using (Font f = new Font("Microsoft JhengHei UI", 24, FontStyle.Bold)) {
                SizeF size = e.Graphics.MeasureString(overlayText, f);
                textRect.Width = size.Width + 20;
                textRect.Height = size.Height + 20;

                // 半透明底色
                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(bgOpacity, bgColor))) {
                    e.Graphics.FillRectangle(bgBrush, textRect);
                }
                // 外框
                using (Pen borderPen = new Pen(borderColor, 3)) {
                    e.Graphics.DrawRectangle(borderPen, textRect.X, textRect.Y, textRect.Width, textRect.Height);
                }
                // 文字
                using (SolidBrush textBrush = new SolidBrush(textColor)) {
                    e.Graphics.DrawString(overlayText, f, textBrush, textRect.X + 10, textRect.Y + 10);
                }
            }
        }

        private Rectangle GetDisplayRect() {
            if (baseCollage == null) return Rectangle.Empty;
            float ratio = Math.Min((float)pb.Width / baseCollage.Width, (float)pb.Height / baseCollage.Height);
            int w = (int)(baseCollage.Width * ratio);
            int h = (int)(baseCollage.Height * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            if (textRect.Contains(e.Location)) {
                isDraggingText = true;
                dragOffset = new PointF(e.X - textRect.X, e.Y - textRect.Y);
            }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            if (isDraggingText) {
                textRect.X = e.X - dragOffset.X;
                textRect.Y = e.Y - dragOffset.Y;
                pb.Invalidate();
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            isDraggingText = false;
        }

        private void SaveImage() {
            if (baseCollage == null) return;

            // 建立最終輸出的高解析度圖片
            Bitmap finalImage = new Bitmap(baseCollage.Width, baseCollage.Height);
            using (Graphics g = Graphics.FromImage(finalImage)) {
                g.DrawImage(baseCollage, 0, 0);

                // 將畫面上的文字座標，換算回高解析度底圖的座標
                Rectangle dispRect = GetDisplayRect();
                float scaleX = (float)baseCollage.Width / dispRect.Width;
                float scaleY = (float)baseCollage.Height / dispRect.Height;
                
                float actualX = (textRect.X - dispRect.X) * scaleX;
                float actualY = (textRect.Y - dispRect.Y) * scaleY;
                float actualW = textRect.Width * scaleX;
                float actualH = textRect.Height * scaleY;

                using (Font f = new Font("Microsoft JhengHei UI", 24 * scaleY, FontStyle.Bold)) {
                    using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(bgOpacity, bgColor))) {
                        g.FillRectangle(bgBrush, actualX, actualY, actualW, actualH);
                    }
                    using (Pen borderPen = new Pen(borderColor, 3 * scaleX)) {
                        g.DrawRectangle(borderPen, actualX, actualY, actualW, actualH);
                    }
                    using (SolidBrush textBrush = new SolidBrush(textColor)) {
                        g.DrawString(overlayText, f, textBrush, actualX + (10 * scaleX), actualY + (10 * scaleY));
                    }
                }
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "Collage_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    finalImage.Save(sfd.FileName, ImageFormat.Jpeg);
                    MessageBox.Show("拼貼儲存成功！");
                }
            }
        }
    }
}
