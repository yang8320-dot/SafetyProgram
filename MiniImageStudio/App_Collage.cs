/* * 功能：進階互動式拼貼模組 (支援拖曳載入、獨立縮放/旋轉、間距調整、自訂文字、全部清除)
 * 對應選單名稱：拼貼
 * 對應資料表名稱：App_Collage
 */
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq; 

namespace MiniImageStudio {
    public class App_Collage : UserControl {
        private class CollageFrame {
            public RectangleF NormalizedRect; 
            public Image Img = null;
            public float Scale = 1.0f;
            public float Angle = 0f;
            public float OffsetX = 0f, OffsetY = 0f;
        }

        private PictureBox pb;
        private List<CollageFrame> frames = new List<CollageFrame>();
        private int activeFrameIndex = -1;
        private int spacing = 10; 
        private int baseCanvasSize = 1500; 

        private ComboBox cbLayout;
        private TrackBar tbSpacing, tbScale, tbRotate;
        private Label lblActiveFrame;

        private bool isTextModeActive = false;
        private string textContent = "拖曳此文字";
        private Font textFont = new Font("Microsoft JhengHei UI", 36, FontStyle.Bold);
        private Color textColor = Color.White, textBgColor = Color.Black, textBorderColor = Color.White;
        private int textOpacity = 150;
        private RectangleF textRect = new RectangleF(50, 50, 250, 80);
        
        private bool isDraggingText = false, isPanningImage = false;
        private PointF dragOffset;
        private Point lastMousePos;

        public App_Collage() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
            LoadTemplate(0); 
        }

        private void InitializeUI() {
            // 面板高度 130，完美容納兩排按鈕
            Panel ctrlPanel = new Panel { Dock = DockStyle.Top, Height = 130, BackColor = SystemColors.Control };
            
            // ================== 第一排 (Top: 15) 全域控制與選取框控制 ==================
            Label lblLayout = new Label { Text = "模版:", Left = 15, Top = 25, AutoSize = true };
            cbLayout = new ComboBox { Left = 65, Top = 20, Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
            cbLayout.Items.AddRange(new string[] { "上下兩張", "左右兩張", "上1 下2", "左1 右2", "左2 右1" });
            cbLayout.SelectedIndex = 0;

            Label lblSpacing = new Label { Text = "間距:", Left = 185, Top = 25, AutoSize = true };
            tbSpacing = new TrackBar { Left = 230, Top = 20, Width = 80, Minimum = 0, Maximum = 100, Value = spacing, TickStyle = TickStyle.None };

            Label lblScale = new Label { Text = "縮放:", Left = 315, Top = 25, AutoSize = true, ForeColor = Color.Blue };
            tbScale = new TrackBar { Left = 360, Top = 20, Width = 80, Minimum = 10, Maximum = 300, Value = 100, TickStyle = TickStyle.None, Enabled = false };
            
            Label lblRotate = new Label { Text = "旋轉:", Left = 445, Top = 25, AutoSize = true, ForeColor = Color.Blue };
            tbRotate = new TrackBar { Left = 490, Top = 20, Width = 80, Minimum = -180, Maximum = 180, Value = 0, TickStyle = TickStyle.None, Enabled = false };

            Button btnClearFrame = new Button { Text = "清除所選", Left = 580, Top = 15, Width = 80, Height = 40 };
            Button btnClearAll = new Button { Text = "全部清除", Left = 670, Top = 15, Width = 80, Height = 40, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存拼貼圖", Left = 760, Top = 15, Width = 100, Height = 40, BackColor = Color.SeaGreen, ForeColor = Color.White };

            // ================== 第二排 (Top: 75) 插入文字控制 ==================
            Button btnInsertText = new Button { Text = "插入文字", Left = 15, Top = 75, Width = 90, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White };
            TextBox txtInput = new TextBox { Left = 115, Top = 82, Width = 130, Text = textContent };
            
            Button btnFont = new Button { Text = "字體", Left = 255, Top = 75, Width = 55, Height = 40 };
            Button btnTextColor = new Button { Text = "字色", Left = 320, Top = 75, Width = 55, Height = 40, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Left = 385, Top = 75, Width = 55, Height = 40, BackColor = textBgColor };
            
            Label lblOpacity = new Label { Text = "透明度:", Left = 450, Top = 85, AutoSize = true };
            TrackBar tbOpacity = new TrackBar { Left = 510, Top = 75, Width = 100, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };
            lblActiveFrame = new Label { Text = "紅框代表目前選取的圖片", Left = 620, Top = 85, AutoSize = true, ForeColor = Color.Red };

            // --- 綁定事件 ---
            cbLayout.SelectedIndexChanged += (s, e) => LoadTemplate(cbLayout.SelectedIndex);
            tbSpacing.ValueChanged += (s, e) => { spacing = tbSpacing.Value; pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();

            tbScale.ValueChanged += (s, e) => { if (activeFrameIndex >= 0) { frames[activeFrameIndex].Scale = tbScale.Value / 100f; pb.Invalidate(); } };
            tbRotate.ValueChanged += (s, e) => { if (activeFrameIndex >= 0) { frames[activeFrameIndex].Angle = tbRotate.Value; pb.Invalidate(); } };
            
            btnClearFrame.Click += (s, e) => {
                if (activeFrameIndex >= 0 && frames[activeFrameIndex].Img != null) {
                    frames[activeFrameIndex].Img.Dispose(); frames[activeFrameIndex].Img = null; pb.Invalidate();
                }
            };
            btnClearAll.Click += (s, e) => {
                foreach (var f in frames) if (f.Img != null) f.Img.Dispose();
                LoadTemplate(cbLayout.SelectedIndex); // 重新載入空白模版
                isTextModeActive = false; pb.Invalidate();
            };

            btnInsertText.Click += (s, e) => { isTextModeActive = true; pb.Invalidate(); };
            txtInput.TextChanged += (s, e) => { textContent = txtInput.Text; if (isTextModeActive) pb.Invalidate(); };
            btnFont.Click += (s, e) => { using (FontDialog fd = new FontDialog { Font = textFont }) { if (fd.ShowDialog() == DialogResult.OK) { textFont = fd.Font; pb.Invalidate(); } } };
            btnTextColor.Click += (s, e) => ChooseColor(ref textColor, btnTextColor);
            btnBgColor.Click += (s, e) => ChooseColor(ref textBgColor, btnBgColor);
            tbOpacity.ValueChanged += (s, e) => { textOpacity = tbOpacity.Value; pb.Invalidate(); };

            ctrlPanel.Controls.AddRange(new Control[] { 
                lblLayout, cbLayout, lblSpacing, tbSpacing, lblScale, tbScale, lblRotate, tbRotate, btnClearFrame, btnClearAll, btnSave,
                btnInsertText, txtInput, btnFont, btnTextColor, btnBgColor, lblOpacity, tbOpacity, lblActiveFrame
            });

            // --- 圖片預覽區 ---
            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, Cursor = Cursors.Cross, AllowDrop = true };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;
            pb.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            pb.DragDrop += Pb_DragDrop;

            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15 });
            this.Controls.Add(ctrlPanel);
        }

        private void ChooseColor(ref Color target, Button btn) {
            using (ColorDialog cd = new ColorDialog { Color = target }) {
                if (cd.ShowDialog() == DialogResult.OK) { target = cd.Color; btn.BackColor = target; pb.Invalidate(); }
            }
        }

        private void LoadTemplate(int index) {
            List<Image> oldImages = new List<Image>();
            foreach (var f in frames) if (f.Img != null) oldImages.Add(f.Img);

            frames.Clear(); activeFrameIndex = -1; UpdateActiveFrameUI();

            if (index == 0) { frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 1, 0.5f) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0.5f, 1, 0.5f) }); } 
            else if (index == 1) { frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 0.5f, 1) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0, 0.5f, 1) }); } 
            else if (index == 2) { frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 1, 0.5f) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0.5f, 0.5f, 0.5f) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0.5f, 0.5f, 0.5f) }); } 
            else if (index == 3) { frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 0.5f, 1) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0, 0.5f, 0.5f) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0.5f, 0.5f, 0.5f) }); } 
            else if (index == 4) { frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 0.5f, 0.5f) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0.5f, 0.5f, 0.5f) }); frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0, 0.5f, 1) }); }

            for (int i = 0; i < Math.Min(oldImages.Count, frames.Count); i++) {
                frames[i].Img = oldImages[i]; ResetFrameTransform(frames[i]);
            }
            pb.Invalidate();
        }

        private void ResetFrameTransform(CollageFrame frame) {
            frame.Scale = 1.0f; frame.Angle = 0f; frame.OffsetX = 0f; frame.OffsetY = 0f;
            if (frame.Img != null) {
                RectangleF actualRect = GetActualFrameRect(frame, baseCanvasSize, baseCanvasSize);
                frame.Scale = Math.Max(actualRect.Width / frame.Img.Width, actualRect.Height / frame.Img.Height);
            }
        }

        private Rectangle GetDisplayRect() {
            float ratio = Math.Min((float)pb.Width / baseCanvasSize, (float)pb.Height / baseCanvasSize);
            int w = (int)(baseCanvasSize * ratio), h = (int)(baseCanvasSize * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private RectangleF GetActualFrameRect(CollageFrame frame, int W, int H) {
            float x = frame.NormalizedRect.X * W + spacing;
            float y = frame.NormalizedRect.Y * H + spacing;
            float w = frame.NormalizedRect.Width * W - (spacing * 2);
            float h = frame.NormalizedRect.Height * H - (spacing * 2);
            return new RectangleF(x, y, w, h);
        }

        private void Pb_Paint(object sender, PaintEventArgs e) {
            Rectangle dispRect = GetDisplayRect();
            e.Graphics.FillRectangle(Brushes.White, dispRect); 

            float scaleX = (float)dispRect.Width / baseCanvasSize, scaleY = (float)dispRect.Height / baseCanvasSize;

            for (int i = 0; i < frames.Count; i++) {
                var frame = frames[i];
                RectangleF highResRect = GetActualFrameRect(frame, baseCanvasSize, baseCanvasSize);
                RectangleF drawRect = new RectangleF(dispRect.X + highResRect.X * scaleX, dispRect.Y + highResRect.Y * scaleY, highResRect.Width * scaleX, highResRect.Height * scaleY);

                e.Graphics.SetClip(drawRect);
                if (frame.Img != null) {
                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    Matrix m = new Matrix();
                    PointF center = new PointF(drawRect.X + drawRect.Width / 2 + (frame.OffsetX * scaleX), drawRect.Y + drawRect.Height / 2 + (frame.OffsetY * scaleY));
                    m.Translate(center.X, center.Y); m.Rotate(frame.Angle); m.Scale(frame.Scale * scaleX, frame.Scale * scaleY); m.Translate(-frame.Img.Width / 2f, -frame.Img.Height / 2f);
                    e.Graphics.Transform = m; e.Graphics.DrawImage(frame.Img, Point.Empty); e.Graphics.ResetTransform();
                } else {
                    using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center }) {
                        e.Graphics.DrawString("點擊或拖曳圖片", MainForm.UI_Font, Brushes.Gray, drawRect, sf);
                    }
                }
                e.Graphics.ResetClip();

                using (Pen borderPen = new Pen(i == activeFrameIndex ? Color.Red : Color.LightGray, i == activeFrameIndex ? 3 : 1)) {
                    e.Graphics.DrawRectangle(borderPen, drawRect.X, drawRect.Y, drawRect.Width, drawRect.Height);
                }
            }

            if (isTextModeActive && !string.IsNullOrEmpty(textContent)) {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                SizeF size = e.Graphics.MeasureString(textContent, textFont);
                textRect.Width = size.Width + 20; textRect.Height = size.Height + 20;

                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(textOpacity, textBgColor))) e.Graphics.FillRectangle(bgBrush, textRect);
                using (Pen borderPen = new Pen(textBorderColor, 3)) e.Graphics.DrawRectangle(borderPen, textRect.X, textRect.Y, textRect.Width, textRect.Height);
                using (SolidBrush tBrush = new SolidBrush(textColor)) e.Graphics.DrawString(textContent, textFont, tBrush, textRect.X + 10, textRect.Y + 10);
            }
        }

        private int GetFrameIndexAtPoint(Point p) {
            Rectangle dispRect = GetDisplayRect();
            float scaleX = (float)baseCanvasSize / dispRect.Width, scaleY = (float)baseCanvasSize / dispRect.Height;
            float x = (p.X - dispRect.X) * scaleX, y = (p.Y - dispRect.Y) * scaleY;
            for (int i = 0; i < frames.Count; i++) {
                RectangleF rect = GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize);
                if (rect.Contains(x, y)) return i;
            }
            return -1;
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            if (isTextModeActive && textRect.Contains(e.Location)) { isDraggingText = true; dragOffset = new PointF(e.X - textRect.X, e.Y - textRect.Y); return; }

            int clickedIndex = GetFrameIndexAtPoint(e.Location);
            if (clickedIndex >= 0) {
                activeFrameIndex = clickedIndex; UpdateActiveFrameUI(); pb.Invalidate();
                if (frames[clickedIndex].Img == null) {
                    using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Images|*.jpg;*.png;*.bmp" }) {
                        if (ofd.ShowDialog() == DialogResult.OK) { frames[clickedIndex].Img = Image.FromFile(ofd.FileName); ResetFrameTransform(frames[clickedIndex]); pb.Invalidate(); }
                    }
                } else { isPanningImage = true; lastMousePos = e.Location; }
            } else { activeFrameIndex = -1; UpdateActiveFrameUI(); pb.Invalidate(); }
        }

        private void UpdateActiveFrameUI() {
            if (activeFrameIndex >= 0) {
                tbScale.Enabled = true; tbRotate.Enabled = true;
                tbScale.Value = (int)(frames[activeFrameIndex].Scale * 100);
                tbRotate.Value = (int)frames[activeFrameIndex].Angle;
            } else {
                tbScale.Enabled = false; tbRotate.Enabled = false;
                tbScale.Value = 100; tbRotate.Value = 0;
            }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            if (isDraggingText) { textRect.X = e.X - dragOffset.X; textRect.Y = e.Y - dragOffset.Y; pb.Invalidate(); } 
            else if (isPanningImage && activeFrameIndex >= 0) {
                float scaleRatio = (float)baseCanvasSize / GetDisplayRect().Width;
                frames[activeFrameIndex].OffsetX += (e.X - lastMousePos.X) * scaleRatio;
                frames[activeFrameIndex].OffsetY += (e.Y - lastMousePos.Y) * scaleRatio;
                lastMousePos = e.Location; pb.Invalidate();
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) { isDraggingText = false; isPanningImage = false; }

        private void Pb_DragDrop(object sender, DragEventArgs e) {
            Point clientPoint = pb.PointToClient(new Point(e.X, e.Y));
            int dropIndex = GetFrameIndexAtPoint(clientPoint);
            if (dropIndex >= 0) {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].ToLower().EndsWith(".jpg") || files[0].ToLower().EndsWith(".png"))) {
                    if (frames[dropIndex].Img != null) frames[dropIndex].Img.Dispose();
                    frames[dropIndex].Img = Image.FromFile(files[0]);
                    ResetFrameTransform(frames[dropIndex]); activeFrameIndex = dropIndex;
                    UpdateActiveFrameUI(); pb.Invalidate();
                }
            }
        }

        private void SaveImage() {
            Bitmap finalImg = new Bitmap(baseCanvasSize, baseCanvasSize);
            using (Graphics g = Graphics.FromImage(finalImg)) {
                g.FillRectangle(Brushes.White, 0, 0, baseCanvasSize, baseCanvasSize);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                foreach (var frame in frames) {
                    RectangleF highResRect = GetActualFrameRect(frame, baseCanvasSize, baseCanvasSize);
                    g.SetClip(highResRect);
                    if (frame.Img != null) {
                        Matrix m = new Matrix();
                        PointF center = new PointF(highResRect.X + highResRect.Width / 2 + frame.OffsetX, highResRect.Y + highResRect.Height / 2 + frame.OffsetY);
                        m.Translate(center.X, center.Y); m.Rotate(frame.Angle); m.Scale(frame.Scale, frame.Scale); m.Translate(-frame.Img.Width / 2f, -frame.Img.Height / 2f);
                        g.Transform = m; g.DrawImage(frame.Img, Point.Empty); g.ResetTransform();
                    }
                    g.ResetClip();
                }

                if (isTextModeActive && !string.IsNullOrEmpty(textContent)) {
                    Rectangle dispRect = GetDisplayRect();
                    float scaleX = (float)baseCanvasSize / dispRect.Width, scaleY = (float)baseCanvasSize / dispRect.Height;
                    float actualX = (textRect.X - dispRect.X) * scaleX, actualY = (textRect.Y - dispRect.Y) * scaleY;
                    using (Font f = new Font(textFont.FontFamily, textFont.Size * scaleY, textFont.Style)) {
                        using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(textOpacity, textBgColor))) g.FillRectangle(bgBrush, actualX, actualY, textRect.Width * scaleX, textRect.Height * scaleY);
                        using (Pen borderPen = new Pen(textBorderColor, 3 * scaleX)) g.DrawRectangle(borderPen, actualX, actualY, textRect.Width * scaleX, textRect.Height * scaleY);
                        using (SolidBrush tBrush = new SolidBrush(textColor)) g.DrawString(textContent, f, tBrush, actualX + (10 * scaleX), actualY + (10 * scaleY));
                    }
                }
            }

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "Collage_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    EncoderParameters ep = new EncoderParameters(1);
                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L); 
                    ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                    finalImg.Save(sfd.FileName, codec, ep); MessageBox.Show("拼貼圖儲存成功！");
                }
            }
        }
    }
}
