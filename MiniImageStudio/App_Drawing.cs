/* * 功能：繪製模組 (修復介面重疊，加入獨立的動態文字框插入功能)
 * 對應選單名稱：繪製
 * 對應資料表名稱：App_Drawing
 */
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class App_Drawing : UserControl {
        private PictureBox pb;
        private Bitmap canvas;
        
        // 繪筆設定
        private Color penColor = Color.Red;
        private string drawMode = "Line";
        private int penWidth = 3;
        private Point startPoint;
        private bool isDrawing = false;

        // 文字框設定
        private bool isTextModeActive = false;
        private string textContent = "請輸入文字";
        private Font textFont = new Font("Microsoft JhengHei UI", 24, FontStyle.Bold);
        private Color textColor = Color.Black;
        private Color textBgColor = Color.Yellow;
        private Color textBorderColor = Color.Red;
        private int textOpacity = 200; // 0-255
        
        private RectangleF textRect = new RectangleF(50, 50, 200, 60);
        private bool isDraggingText = false;
        private PointF dragOffset;

        public App_Drawing() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
        }

        private void InitializeUI() {
            // 控制面板加大到 130 避免重疊
            Panel ctrlPanel = new Panel { Dock = DockStyle.Top, Height = 130, BackColor = SystemColors.Control };
            
            // --- 第一排：繪圖控制 (Top: 10) ---
            Button btnLoad = new Button { Text = "載入圖片", Left = 15, Top = 10, Width = 100, Height = 35 };
            Button btnPenColor = new Button { Text = "畫筆顏色", Left = 125, Top = 10, Width = 100, Height = 35, BackColor = penColor };
            
            ComboBox cbMode = new ComboBox { Left = 235, Top = 15, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cbMode.Items.AddRange(new string[] { "自由畫線", "標準矩形" });
            cbMode.SelectedIndex = 0;

            ComboBox cbSize = new ComboBox { Left = 345, Top = 15, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbSize.Items.AddRange(new string[] { "細(2pt)", "中(5pt)", "粗(10pt)" });
            cbSize.SelectedIndex = 0;

            Button btnSave = new Button { Text = "儲存圖片", Left = 440, Top = 10, Width = 100, Height = 35, BackColor = Color.SeaGreen, ForeColor = Color.White };

            // --- 第二排：插入文字框控制 (Top: 65) ---
            Button btnInsertText = new Button { Text = "插入文字框", Left = 15, Top = 65, Width = 110, Height = 35, BackColor = Color.SteelBlue, ForeColor = Color.White };
            TextBox txtInput = new TextBox { Left = 135, Top = 70, Width = 150, Text = textContent };
            
            Button btnFont = new Button { Text = "字體", Left = 295, Top = 65, Width = 60, Height = 35 };
            Button btnTextColor = new Button { Text = "字色", Left = 360, Top = 65, Width = 60, Height = 35, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Left = 425, Top = 65, Width = 60, Height = 35, BackColor = textBgColor };
            Button btnBorderColor = new Button { Text = "框色", Left = 490, Top = 65, Width = 60, Height = 35, BackColor = textBorderColor };
            
            Label lblOpacity = new Label { Text = "透明度:", Left = 560, Top = 75, AutoSize = true };
            TrackBar tbOpacity = new TrackBar { Left = 620, Top = 65, Width = 120, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };

            Button btnApplyText = new Button { Text = "合併文字與圖片", Left = 750, Top = 65, Width = 140, Height = 35, BackColor = Color.DarkOrange, ForeColor = Color.White };

            // 事件綁定
            btnLoad.Click += (s, e) => LoadImage();
            btnPenColor.Click += (s, e) => ChooseColor(ref penColor, btnPenColor);
            cbMode.SelectedIndexChanged += (s, e) => drawMode = cbMode.SelectedIndex == 0 ? "Line" : "Frame";
            cbSize.SelectedIndexChanged += (s, e) => penWidth = cbSize.SelectedIndex == 0 ? 2 : (cbSize.SelectedIndex == 1 ? 5 : 10);
            btnSave.Click += (s, e) => SaveImage();

            btnInsertText.Click += (s, e) => { isTextModeActive = true; pb.Invalidate(); };
            txtInput.TextChanged += (s, e) => { textContent = txtInput.Text; if (isTextModeActive) pb.Invalidate(); };
            btnFont.Click += (s, e) => {
                using (FontDialog fd = new FontDialog { Font = textFont }) {
                    if (fd.ShowDialog() == DialogResult.OK) { textFont = fd.Font; pb.Invalidate(); }
                }
            };
            btnTextColor.Click += (s, e) => ChooseColor(ref textColor, btnTextColor);
            btnBgColor.Click += (s, e) => ChooseColor(ref textBgColor, btnBgColor);
            btnBorderColor.Click += (s, e) => ChooseColor(ref textBorderColor, btnBorderColor);
            tbOpacity.ValueChanged += (s, e) => { textOpacity = tbOpacity.Value; pb.Invalidate(); };
            
            btnApplyText.Click += (s, e) => ApplyTextToImage();

            ctrlPanel.Controls.AddRange(new Control[] { 
                btnLoad, btnPenColor, cbMode, cbSize, btnSave,
                btnInsertText, txtInput, btnFont, btnTextColor, btnBgColor, btnBorderColor, lblOpacity, tbOpacity, btnApplyText
            });

            pb = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.DarkGray, BorderStyle = BorderStyle.Fixed3D };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;

            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15 });
            this.Controls.Add(ctrlPanel);
        }

        private void ChooseColor(ref Color target, Button btn) {
            using (ColorDialog cd = new ColorDialog { Color = target }) {
                if (cd.ShowDialog() == DialogResult.OK) { target = cd.Color; btn.BackColor = target; pb.Invalidate(); }
            }
        }

        private void LoadImage() {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Image|*.jpg;*.png;*.bmp" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    canvas = new Bitmap(ofd.FileName);
                    pb.Image = canvas;
                }
            }
        }

        private Point TranslatePoint(Point p) {
            if (pb.Image == null) return p;
            float imageAspect = (float)pb.Image.Width / pb.Image.Height;
            float controlAspect = (float)pb.Width / pb.Height;
            float scale, offsetX = 0, offsetY = 0;
            if (imageAspect > controlAspect) {
                scale = (float)pb.Width / pb.Image.Width;
                offsetY = (pb.Height - pb.Image.Height * scale) / 2;
            } else {
                scale = (float)pb.Height / pb.Image.Height;
                offsetX = (pb.Width - pb.Image.Width * scale) / 2;
            }
            return new Point((int)((p.X - offsetX) / scale), (int)((p.Y - offsetY) / scale));
        }

        private Rectangle GetDisplayRect() {
            if (pb.Image == null) return Rectangle.Empty;
            float ratio = Math.Min((float)pb.Width / pb.Image.Width, (float)pb.Height / pb.Image.Height);
            int w = (int)(pb.Image.Width * ratio);
            int h = (int)(pb.Image.Height * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private void Pb_Paint(object sender, PaintEventArgs e) {
            if (isTextModeActive && !string.IsNullOrEmpty(textContent)) {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                SizeF size = e.Graphics.MeasureString(textContent, textFont);
                textRect.Width = size.Width + 20;
                textRect.Height = size.Height + 20;

                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(textOpacity, textBgColor))) {
                    e.Graphics.FillRectangle(bgBrush, textRect);
                }
                using (Pen borderPen = new Pen(textBorderColor, 3)) {
                    e.Graphics.DrawRectangle(borderPen, textRect.X, textRect.Y, textRect.Width, textRect.Height);
                }
                using (SolidBrush textBrush = new SolidBrush(textColor)) {
                    e.Graphics.DrawString(textContent, textFont, textBrush, textRect.X + 10, textRect.Y + 10);
                }
            }
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            if (isTextModeActive && textRect.Contains(e.Location)) {
                isDraggingText = true;
                dragOffset = new PointF(e.X - textRect.X, e.Y - textRect.Y);
                return;
            }

            if (canvas != null && !isTextModeActive) {
                isDrawing = true;
                startPoint = TranslatePoint(e.Location);
            }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            if (isDraggingText) {
                textRect.X = e.X - dragOffset.X;
                textRect.Y = e.Y - dragOffset.Y;
                pb.Invalidate();
            } else if (isDrawing) {
                pb.Invalidate(); // 如果需要預覽畫線過程可在此擴充
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            if (isDraggingText) { isDraggingText = false; return; }
            if (!isDrawing || canvas == null) return;
            
            isDrawing = false;
            Point endPoint = TranslatePoint(e.Location);

            using (Graphics g = Graphics.FromImage(canvas)) {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (Pen p = new Pen(penColor, penWidth)) {
                    if (drawMode == "Line") g.DrawLine(p, startPoint, endPoint);
                    else if (drawMode == "Frame") {
                        int x = Math.Min(startPoint.X, endPoint.X);
                        int y = Math.Min(startPoint.Y, endPoint.Y);
                        g.DrawRectangle(p, x, y, Math.Abs(startPoint.X - endPoint.X), Math.Abs(startPoint.Y - endPoint.Y));
                    }
                }
            }
            pb.Refresh();
        }

        private void ApplyTextToImage() {
            if (!isTextModeActive || canvas == null) return;
            
            using (Graphics g = Graphics.FromImage(canvas)) {
                Rectangle dispRect = GetDisplayRect();
                float scaleX = (float)canvas.Width / dispRect.Width;
                float scaleY = (float)canvas.Height / dispRect.Height;
                
                float actualX = (textRect.X - dispRect.X) * scaleX;
                float actualY = (textRect.Y - dispRect.Y) * scaleY;
                float actualW = textRect.Width * scaleX;
                float actualH = textRect.Height * scaleY;

                using (Font f = new Font(textFont.FontFamily, textFont.Size * scaleY, textFont.Style)) {
                    using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(textOpacity, textBgColor))) {
                        g.FillRectangle(bgBrush, actualX, actualY, actualW, actualH);
                    }
                    using (Pen borderPen = new Pen(textBorderColor, 3 * scaleX)) {
                        g.DrawRectangle(borderPen, actualX, actualY, actualW, actualH);
                    }
                    using (SolidBrush textBrush = new SolidBrush(textColor)) {
                        g.DrawString(textContent, f, textBrush, actualX + (10 * scaleX), actualY + (10 * scaleY));
                    }
                }
            }
            isTextModeActive = false; // 合併後關閉拖曳模式
            pb.Refresh();
            MessageBox.Show("文字已合併至圖片！");
        }

        private void SaveImage() {
            if (canvas == null) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PNG Image|*.png|JPG Image|*.jpg" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    canvas.Save(sfd.FileName);
                    MessageBox.Show("儲存成功！");
                }
            }
        }
    }
}
