/* * 功能：拼貼與繪製終極版 (完美解決 GroupBox 下方間距過大 Bug，採用緊湊固定高度排版)
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
            public float Scale = 1.0f, Angle = 0f, OffsetX = 0f, OffsetY = 0f;
        }

        private PictureBox pb;
        private List<CollageFrame> frames = new List<CollageFrame>();
        private List<App_Drawing.DrawShape> shapes = new List<App_Drawing.DrawShape>();
        private App_Drawing.DrawShape selectedShape = null, drawingShape = null;
        private TextBox editBox = null;
        private App_Drawing.DrawShape editingShape = null;

        private int activeFrameIndex = -1, spacing = 10, baseCanvasSize = 1500; 
        private float div1 = 0.5f, div2 = 0.5f;
        private int activeDivider = 0; 

        // 新增 cbRatio 宣告
        private ComboBox cbRatio, cbLayout, cbAlign, cbMode;
        private TrackBar tbSpacing, tbScale, tbRotate;
        private NumericUpDown numPenSize;
        
        private string drawMode = "Select";
        private Color penColor = Color.Red;
        private bool isTextModeActive = false;
        private string textContent = "請輸入文字...";
        private Font textFont = new Font("Microsoft JhengHei UI", 36, FontStyle.Bold);
        private Color textColor = Color.White, textBgColor = Color.Black, textBorderColor = Color.White;
        private int textOpacity = 150;
        
        private bool isDraggingShape = false, isResizingText = false, isPanningImage = false;
        private Point lastMousePos;

        public App_Collage() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
            LoadTemplate(0); 
        }

        private void InitializeUI() {
            FlowLayoutPanel mainFlow = new FlowLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                FlowDirection = FlowDirection.LeftToRight, 
                BackColor = SystemColors.Control,
                Padding = new Padding(5)
            };

            // ================== 小框 1：模版與版面 ==================
            // 加寬以容納畫布比例選單
            GroupBox gb1 = new GroupBox { Text = "模版與排版", Size = new Size(580, 65), Margin = new Padding(5) };
            
            Label lblRatio = new Label { Text = "畫布比例:", Location = new Point(10, 30), AutoSize = true };
            cbRatio = new ComboBox { Location = new Point(75, 25), Width = 65, DropDownStyle = ComboBoxStyle.DropDownList };
            cbRatio.Items.AddRange(new string[] { "1:1", "4:3", "3:4", "16:9", "9:16" });
            cbRatio.SelectedIndex = 0;

            cbLayout = new ComboBox { Location = new Point(150, 25), Width = 95, DropDownStyle = ComboBoxStyle.DropDownList };
            cbLayout.Items.AddRange(new string[] { "上下兩張", "左右兩張", "上1 下2", "左1 右2", "左2 右1" }); 
            cbLayout.SelectedIndex = 0;
            
            Label lblSpacing = new Label { Text = "間距:", Location = new Point(250, 30), AutoSize = true };
            tbSpacing = new TrackBar { Location = new Point(290, 24), Width = 90, Minimum = 0, Maximum = 100, Value = spacing, TickStyle = TickStyle.None };
            
            Button btnClearAll = new Button { Text = "全部清除", Location = new Point(390, 22), Width = 80, Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存拼貼圖", Location = new Point(480, 22), Width = 90, Height = 30, BackColor = Color.SeaGreen, ForeColor = Color.White };
            
            gb1.Controls.AddRange(new Control[] { lblRatio, cbRatio, cbLayout, lblSpacing, tbSpacing, btnClearAll, btnSave });

            // ================== 小框 2：圖片控制 ==================
            GroupBox gb2 = new GroupBox { Text = "選取圖片控制", Size = new Size(430, 65), Margin = new Padding(5) };
            Label lblScale = new Label { Text = "縮放:", Location = new Point(15, 30), AutoSize = true, ForeColor = Color.Blue };
            tbScale = new TrackBar { Location = new Point(55, 24), Width = 110, Minimum = 10, Maximum = 300, Value = 100, TickStyle = TickStyle.None, Enabled = false };
            Label lblRotate = new Label { Text = "旋轉:", Location = new Point(175, 30), AutoSize = true, ForeColor = Color.Blue };
            tbRotate = new TrackBar { Location = new Point(215, 24), Width = 110, Minimum = -180, Maximum = 180, Value = 0, TickStyle = TickStyle.None, Enabled = false };
            Button btnClearFrame = new Button { Text = "刪除圖片", Location = new Point(335, 22), Width = 80, Height = 30 };
            gb2.Controls.AddRange(new Control[] { lblScale, tbScale, lblRotate, tbRotate, btnClearFrame });

            // ================== 小框 3：文字工具 ==================
            GroupBox gb3 = new GroupBox { Text = "文字工具 (雙擊框可編輯)", Size = new Size(535, 65), Margin = new Padding(5) };
            Button btnInsertText = new Button { Text = "插入文字框", Location = new Point(15, 22), Width = 95, Height = 30, BackColor = Color.SteelBlue, ForeColor = Color.White };
            
            // 【修正對齊】將 cbAlign 的 Y 座標由 26 改為 24
            cbAlign = new ComboBox { Location = new Point(120, 24), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cbAlign.Items.AddRange(new string[] { "靠左", "置中", "靠右" }); 
            cbAlign.SelectedIndex = 0;
            
            Button btnFont = new Button { Text = "字體", Location = new Point(190, 22), Width = 55, Height = 30 };
            Button btnTextColor = new Button { Text = "字色", Location = new Point(255, 22), Width = 55, Height = 30, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Location = new Point(320, 22), Width = 55, Height = 30, BackColor = textBgColor };
            Label lblOpacity = new Label { Text = "透明:", Location = new Point(385, 30), AutoSize = true };
            TrackBar tbOpacity = new TrackBar { Location = new Point(425, 24), Width = 100, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };
            gb3.Controls.AddRange(new Control[] { btnInsertText, cbAlign, btnFont, btnTextColor, btnBgColor, lblOpacity, tbOpacity });

            // ================== 小框 4：繪圖工具 ==================
            GroupBox gb4 = new GroupBox { Text = "繪圖工具", Size = new Size(270, 65), Margin = new Padding(5) };
            cbMode = new ComboBox { Location = new Point(15, 26), Width = 70, DropDownStyle = ComboBoxStyle.DropDownList };
            cbMode.Items.AddRange(new string[] { "選取", "畫框", "畫線", "畫圓" }); 
            cbMode.SelectedIndex = 0;
            numPenSize = new NumericUpDown { Location = new Point(95, 26), Minimum = 1, Maximum = 10, Value = 5, Width = 45 };
            Button btnPenColor = new Button { Location = new Point(150, 24), Width = 30, Height = 28, BackColor = penColor };
            Button btnUndo = new Button { Text = "返回", Location = new Point(190, 22), Width = 65, Height = 30 };
            gb4.Controls.AddRange(new Control[] { cbMode, numPenSize, btnPenColor, btnUndo });

            mainFlow.Controls.AddRange(new Control[] { gb1, gb2, gb3, gb4 });

            // 畫布區
            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.DarkGray, Cursor = Cursors.Default };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;
            pb.DoubleClick += Pb_DoubleClick;

            this.Controls.Add(pb);
            this.Controls.Add(mainFlow);

            // 綁定事件
            cbRatio.SelectedIndexChanged += (s, e) => { pb.Invalidate(); };
            cbLayout.SelectedIndexChanged += (s, e) => LoadTemplate(cbLayout.SelectedIndex);
            tbSpacing.ValueChanged += (s, e) => { spacing = tbSpacing.Value; pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();
            
            tbScale.ValueChanged += (s, e) => { 
                if (activeFrameIndex >= 0) { 
                    frames[activeFrameIndex].Scale = tbScale.Value / 100f; 
                    pb.Invalidate(); 
                } 
            };
            tbRotate.ValueChanged += (s, e) => { 
                if (activeFrameIndex >= 0) { 
                    frames[activeFrameIndex].Angle = tbRotate.Value; 
                    pb.Invalidate(); 
                } 
            };
            btnClearFrame.Click += (s, e) => { 
                if (activeFrameIndex >= 0) { 
                    frames[activeFrameIndex].Img = null; 
                    pb.Invalidate(); 
                } 
            };
            
            btnInsertText.Click += (s, e) => { 
                shapes.Add(new App_Drawing.DrawShape { 
                    Type = "Text", Text = textContent, Font = textFont, Color = textColor, 
                    BgColor = textBgColor, TextAlign = cbAlign.SelectedItem.ToString(), 
                    Opacity = tbOpacity.Value, TextRect = new RectangleF(pb.Width/2 - 100, pb.Height/2 - 50, 200, 100) 
                });
                pb.Invalidate();
            };
            btnFont.Click += (s, e) => { 
                using (FontDialog fd = new FontDialog { Font = textFont }) {
                    if (fd.ShowDialog() == DialogResult.OK) textFont = fd.Font; 
                }
            };
            btnTextColor.Click += (s, e) => { 
                using (ColorDialog cd = new ColorDialog { Color = textColor }) {
                    if (cd.ShowDialog() == DialogResult.OK) { textColor = cd.Color; btnTextColor.BackColor = textColor; } 
                }
            };
            btnBgColor.Click += (s, e) => { 
                using (ColorDialog cd = new ColorDialog { Color = textBgColor }) {
                    if (cd.ShowDialog() == DialogResult.OK) { textBgColor = cd.Color; btnBgColor.BackColor = textBgColor; } 
                }
            };
            
            cbMode.SelectedIndexChanged += (s, e) => {
                if (cbMode.SelectedItem.ToString() == "選取") drawMode = "Select";
                else if (cbMode.SelectedItem.ToString() == "畫框") drawMode = "Frame";
                else if (cbMode.SelectedItem.ToString() == "畫線") drawMode = "Line";
                else if (cbMode.SelectedItem.ToString() == "畫圓") drawMode = "Circle";
            };
            btnPenColor.Click += (s, e) => { 
                using (ColorDialog cd = new ColorDialog { Color = penColor }) {
                    if (cd.ShowDialog() == DialogResult.OK) { penColor = cd.Color; btnPenColor.BackColor = penColor; } 
                }
            };
            btnUndo.Click += (s, e) => { 
                if (shapes.Count > 0) { shapes.RemoveAt(shapes.Count - 1); pb.Invalidate(); } 
            };
        }

        private void LoadTemplate(int index) {
            frames.Clear();
            if (index == 0) { 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 1, 0.5f) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0.5f, 1, 0.5f) }); 
            } else if (index == 1) { 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 0.5f, 1) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0, 0.5f, 1) }); 
            } else if (index == 2) { 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 1, 0.5f) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0.5f, 0.5f, 0.5f) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0.5f, 0.5f, 0.5f) }); 
            } else if (index == 3) { 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 0.5f, 1) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0, 0.5f, 0.5f) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0.5f, 0.5f, 0.5f) }); 
            } else if (index == 4) { 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0, 0.5f, 0.5f) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0, 0.5f, 0.5f, 0.5f) }); 
                frames.Add(new CollageFrame { NormalizedRect = new RectangleF(0.5f, 0, 0.5f, 1) }); 
            }
            pb.Invalidate();
        }

        // 讀取並套用選擇的畫布比例 (已加入防崩潰保護)
        private Rectangle GetCanvasRect() {
            // 加入 Math.Max，確保當視窗極小或尚未載入時，長寬至少為 1，防止算出負數
            int targetWidth = Math.Max(1, pb.Width - 40);
            int targetHeight = Math.Max(1, pb.Height - 40);
            
            string ratioStr = cbRatio != null && cbRatio.SelectedItem != null ? cbRatio.SelectedItem.ToString() : "1:1";
            float ratio = 1.0f;
            if (ratioStr == "4:3") ratio = 4.0f / 3.0f;
            else if (ratioStr == "3:4") ratio = 3.0f / 4.0f;
            else if (ratioStr == "16:9") ratio = 16.0f / 9.0f;
            else if (ratioStr == "9:16") ratio = 9.0f / 16.0f;

            if (targetWidth / (float)targetHeight > ratio) {
                targetWidth = Math.Max(1, (int)(targetHeight * ratio));
            } else {
                targetHeight = Math.Max(1, (int)(targetWidth / ratio));
            }

            return new Rectangle((pb.Width - targetWidth) / 2, (pb.Height - targetHeight) / 2, targetWidth, targetHeight);
        }

        private RectangleF GetFrameRect(CollageFrame f, Rectangle canvas) {
            return new RectangleF(
                canvas.X + f.NormalizedRect.X * canvas.Width + spacing,
                canvas.Y + f.NormalizedRect.Y * canvas.Height + spacing,
                f.NormalizedRect.Width * canvas.Width - spacing * 2,
                f.NormalizedRect.Height * canvas.Height - spacing * 2
            );
        }

        private void Pb_Paint(object sender, PaintEventArgs e) {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle canvasRect = GetCanvasRect();
            e.Graphics.FillRectangle(Brushes.White, canvasRect);

            for (int i = 0; i < frames.Count; i++) {
                RectangleF rect = GetFrameRect(frames[i], canvasRect);
                if (frames[i].Img != null) {
                    e.Graphics.SetClip(rect);
                    e.Graphics.TranslateTransform(rect.X + rect.Width / 2 + frames[i].OffsetX, rect.Y + rect.Height / 2 + frames[i].OffsetY);
                    e.Graphics.RotateTransform(frames[i].Angle);
                    e.Graphics.ScaleTransform(frames[i].Scale, frames[i].Scale);
                    e.Graphics.DrawImage(frames[i].Img, -frames[i].Img.Width / 2, -frames[i].Img.Height / 2);
                    e.Graphics.ResetTransform();
                    e.Graphics.ResetClip();
                } else {
                    e.Graphics.FillRectangle(Brushes.LightGray, rect);
                    StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    e.Graphics.DrawString("點擊加入圖片", this.Font, Brushes.Gray, rect, sf);
                }
                if (i == activeFrameIndex) {
                    e.Graphics.DrawRectangle(Pens.Blue, Rectangle.Round(rect));
                }
            }

            foreach (var s in shapes) {
                if (s.Type == "Text") {
                    using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) {
                        e.Graphics.FillRectangle(bg, s.TextRect);
                    }
                    StringFormat sf = new StringFormat { LineAlignment = StringAlignment.Near };
                    if (s.TextAlign == "置中") sf.Alignment = StringAlignment.Center; 
                    else if (s.TextAlign == "靠右") sf.Alignment = StringAlignment.Far; 
                    else sf.Alignment = StringAlignment.Near;

                    using (SolidBrush tb = new SolidBrush(s.Color)) {
                        e.Graphics.DrawString(s.Text, s.Font, tb, new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20), sf);
                    }
                } else {
                    using (Pen p = new Pen(s.Color, s.PenWidth)) {
                        int x = Math.Min(s.Start.X, s.End.X);
                        int y = Math.Min(s.Start.Y, s.End.Y);
                        int w = Math.Abs(s.Start.X - s.End.X);
                        int h = Math.Abs(s.Start.Y - s.End.Y);
                        if (s.Type == "Line") e.Graphics.DrawLine(p, s.Start, s.End);
                        else if (s.Type == "Frame") e.Graphics.DrawRectangle(p, x, y, w, h);
                        else if (s.Type == "Circle") e.Graphics.DrawEllipse(p, x, y, w, h);
                    }
                }
            }

            if (drawingShape != null) {
                using (Pen p = new Pen(drawingShape.Color, drawingShape.PenWidth)) {
                    int x = Math.Min(drawingShape.Start.X, drawingShape.End.X);
                    int y = Math.Min(drawingShape.Start.Y, drawingShape.End.Y);
                    int w = Math.Abs(drawingShape.Start.X - drawingShape.End.X);
                    int h = Math.Abs(drawingShape.Start.Y - drawingShape.End.Y);
                    if (drawingShape.Type == "Line") e.Graphics.DrawLine(p, drawingShape.Start, drawingShape.End);
                    else if (drawingShape.Type == "Frame") e.Graphics.DrawRectangle(p, x, y, w, h);
                    else if (drawingShape.Type == "Circle") e.Graphics.DrawEllipse(p, x, y, w, h);
                }
            }
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            CommitTextEdit();
            lastMousePos = e.Location;
            
            if (drawMode == "Select") {
                selectedShape = shapes.LastOrDefault(s => s.GetBounds().Contains(e.Location));
                if (selectedShape != null) { 
                    isDraggingShape = true; 
                    return; 
                }

                Rectangle canvasRect = GetCanvasRect();
                for (int i = 0; i < frames.Count; i++) {
                    if (GetFrameRect(frames[i], canvasRect).Contains(e.Location)) {
                        activeFrameIndex = i;
                        if (frames[i].Img == null) {
                            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp" }) {
                                if (ofd.ShowDialog() == DialogResult.OK) { frames[i].Img = Image.FromFile(ofd.FileName); }
                            }
                        } else {
                            isPanningImage = true;
                            tbScale.Enabled = true; 
                            tbRotate.Enabled = true;
                            tbScale.Value = (int)(frames[i].Scale * 100);
                            tbRotate.Value = (int)frames[i].Angle;
                        }
                        pb.Invalidate();
                        return;
                    }
                }
                activeFrameIndex = -1;
                tbScale.Enabled = false; 
                tbRotate.Enabled = false;
                pb.Invalidate();
            } else {
                drawingShape = new App_Drawing.DrawShape { Type = drawMode, Color = penColor, PenWidth = (int)numPenSize.Value, Start = e.Location, End = e.Location };
            }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            if (isDraggingShape && selectedShape != null) {
                selectedShape.Move(e.X - lastMousePos.X, e.Y - lastMousePos.Y);
                lastMousePos = e.Location;
                pb.Invalidate();
            } else if (isPanningImage && activeFrameIndex >= 0) {
                frames[activeFrameIndex].OffsetX += (e.X - lastMousePos.X);
                frames[activeFrameIndex].OffsetY += (e.Y - lastMousePos.Y);
                lastMousePos = e.Location;
                pb.Invalidate();
            } else if (drawingShape != null) {
                drawingShape.End = e.Location;
                pb.Invalidate();
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            isDraggingShape = false;
            isPanningImage = false;
            if (drawingShape != null) {
                shapes.Add(drawingShape);
                drawingShape = null;
                pb.Invalidate();
            }
        }

        private void Pb_DoubleClick(object sender, EventArgs e) {
            Point pt = pb.PointToClient(Cursor.Position);
            var shape = shapes.LastOrDefault(s => s.Type == "Text" && s.GetBounds().Contains(pt));
            if (shape != null) {
                editingShape = shape;
                editBox = new TextBox { 
                    Multiline = true, Text = shape.Text, Font = shape.Font, 
                    Location = Point.Round(shape.TextRect.Location), Size = Size.Round(shape.TextRect.Size) 
                };
                editBox.LostFocus += (s2, e2) => CommitTextEdit();
                pb.Controls.Add(editBox);
                editBox.BringToFront();
                editBox.Focus();
            }
        }

        private void CommitTextEdit() {
            if (editBox != null && editingShape != null) {
                editingShape.Text = editBox.Text;
                pb.Controls.Remove(editBox);
                editBox.Dispose();
                editBox = null;
                editingShape = null;
                pb.Invalidate();
            }
        }

        private void SaveImage() {
            CommitTextEdit();
            Rectangle canvasRect = GetCanvasRect();
            Bitmap bmp = new Bitmap(canvasRect.Width, canvasRect.Height);
            using (Graphics g = Graphics.FromImage(bmp)) {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TranslateTransform(-canvasRect.X, -canvasRect.Y);
                PaintEventArgs pe = new PaintEventArgs(g, canvasRect);
                Pb_Paint(this, pe);
            }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "Collage_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    EncoderParameters ep = new EncoderParameters(1);
                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);
                    ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                    bmp.Save(sfd.FileName, codec, ep);
                    MessageBox.Show("儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
