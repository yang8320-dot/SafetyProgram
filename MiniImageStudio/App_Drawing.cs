/* * 功能：進階向量繪製模組 (完美緊湊排版、支援拖曳上傳、預設提示文字)
 */
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging; // 修正缺少此命名空間導致的 EncoderParameters 錯誤
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using System.Linq; // 修正缺少此命名空間導致的 .First() 錯誤

namespace MiniImageStudio {
    public class App_Drawing : UserControl {
        public class DrawShape {
            public string Type; 
            public Color Color;
            public int PenWidth;
            public Point Start, End;
            
            public string Text;
            public Font Font;
            public Color BgColor, BorderColor;
            public string TextAlign; 
            public int Opacity;
            public RectangleF TextRect;
            public bool IsSelected = false;

            public Rectangle GetBounds() {
                if (Type == "Text") return Rectangle.Round(TextRect);
                int x = Math.Min(Start.X, End.X) - PenWidth - 5;
                int y = Math.Min(Start.Y, End.Y) - PenWidth - 5;
                int w = Math.Abs(Start.X - End.X) + PenWidth * 2 + 10;
                int h = Math.Abs(Start.Y - End.Y) + PenWidth * 2 + 10;
                return new Rectangle(x, y, w, h);
            }

            public void Move(int dx, int dy) {
                if (Type == "Text") { 
                    TextRect.X += dx; 
                    TextRect.Y += dy; 
                } else { 
                    Start.X += dx; Start.Y += dy; End.X += dx; End.Y += dy; 
                }
            }
            
            public RectangleF GetResizeHandle() {
                if (Type == "Text") return new RectangleF(TextRect.Right - 15, TextRect.Bottom - 15, 15, 15);
                return RectangleF.Empty;
            }
        }

        private PictureBox pb;
        private Bitmap canvas;
        private List<DrawShape> shapes = new List<DrawShape>();
        private DrawShape selectedShape = null;
        private DrawShape drawingShape = null;
        private TextBox editBox = null;
        private DrawShape editingShape = null;
        
        private Color penColor = Color.Red;
        private string drawMode = "Frame"; 
        
        private bool isDraggingShape = false, isResizingText = false;
        private Point lastMousePos;
        private bool isTextModeActive = false;
        
        private string textContent = "請輸入文字...";
        private Font textFont = new Font("Microsoft JhengHei UI", 24, FontStyle.Bold);
        private Color textColor = Color.White, textBgColor = Color.Black, textBorderColor = Color.White;
        private int textOpacity = 150;

        private ComboBox cbMode, cbAlign;
        private NumericUpDown numPenSize;

        public App_Drawing() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
        }

        private void InitializeUI() {
            FlowLayoutPanel mainFlow = new FlowLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                FlowDirection = FlowDirection.LeftToRight, 
                BackColor = SystemColors.Control,
                Padding = new Padding(5)
            };

            // ================== 群組 1: 畫布控制 ==================
            GroupBox gb1 = new GroupBox { Text = "畫布", Size = new Size(390, 65), Margin = new Padding(5) };
            Button btnLoad = new Button { Text = "載入圖片", Location = new Point(15, 22), Width = 85, Height = 30 };
            Button btnRotate = new Button { Text = "旋轉圖片", Location = new Point(105, 22), Width = 85, Height = 30 };
            Button btnClear = new Button { Text = "清除全部", Location = new Point(195, 22), Width = 85, Height = 30, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存圖片", Location = new Point(285, 22), Width = 85, Height = 30, BackColor = Color.SeaGreen, ForeColor = Color.White };
            gb1.Controls.AddRange(new Control[] { btnLoad, btnRotate, btnClear, btnSave });

            // ================== 群組 2: 繪圖工具 ==================
            GroupBox gb2 = new GroupBox { Text = "繪圖工具", Size = new Size(265, 65), Margin = new Padding(5) };
            cbMode = new ComboBox { Location = new Point(15, 26), Width = 70, DropDownStyle = ComboBoxStyle.DropDownList };
            cbMode.Items.AddRange(new string[] { "選取", "畫框", "畫線", "畫圓" }); 
            cbMode.SelectedIndex = 1; // 預設畫框
            numPenSize = new NumericUpDown { Location = new Point(95, 26), Minimum = 1, Maximum = 10, Value = 5, Width = 45 };
            Button btnPenColor = new Button { Location = new Point(150, 24), Width = 30, Height = 28, BackColor = penColor }; 
            Button btnUndo = new Button { Text = "返回", Location = new Point(190, 22), Width = 60, Height = 30 };
            gb2.Controls.AddRange(new Control[] { cbMode, numPenSize, btnPenColor, btnUndo });

            // ================== 群組 3: 文字工具 ==================
            GroupBox gb3 = new GroupBox { Text = "文字工具 (雙擊文字框可編輯)", Size = new Size(610, 65), Margin = new Padding(5) };
            Button btnInsertText = new Button { Text = "插入文字框", Location = new Point(15, 22), Width = 95, Height = 30, BackColor = Color.SteelBlue, ForeColor = Color.White };
            
            // 【修正對齊】將 cbAlign 的 Y 座標由 26 改為 24
            cbAlign = new ComboBox { Location = new Point(120, 24), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cbAlign.Items.AddRange(new string[] { "靠左", "置中", "靠右" }); 
            cbAlign.SelectedIndex = 0;
            
            Button btnFont = new Button { Text = "字體", Location = new Point(190, 22), Width = 55, Height = 30 };
            Button btnTextColor = new Button { Text = "字色", Location = new Point(255, 22), Width = 55, Height = 30, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Location = new Point(320, 22), Width = 55, Height = 30, BackColor = textBgColor };
            Label lblOpacity = new Label { Text = "透明度:", Location = new Point(385, 30), AutoSize = true };
            TrackBar tbOpacity = new TrackBar { Location = new Point(445, 24), Width = 150, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };
            gb3.Controls.AddRange(new Control[] { btnInsertText, cbAlign, btnFont, btnTextColor, btnBgColor, lblOpacity, tbOpacity });

            mainFlow.Controls.AddRange(new Control[] { gb1, gb2, gb3 });

            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.DarkGray, Cursor = Cursors.Cross, AllowDrop = true };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;
            pb.DoubleClick += Pb_DoubleClick;
            
            // 支援拖曳圖片
            pb.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            pb.DragDrop += (s, e) => {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0) LoadImageToCanvas(files[0]);
            };

            this.Controls.Add(pb);
            this.Controls.Add(mainFlow);

            // 事件綁定
            btnLoad.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) LoadImageToCanvas(ofd.FileName);
                }
            };
            btnRotate.Click += (s, e) => { if (canvas != null) { canvas.RotateFlip(RotateFlipType.Rotate90FlipNone); pb.Invalidate(); } };
            btnClear.Click += (s, e) => { shapes.Clear(); pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();
            
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

            btnInsertText.Click += (s, e) => { 
                shapes.Add(new DrawShape { 
                    Type = "Text", Text = textContent, Font = textFont, Color = textColor, 
                    BgColor = textBgColor, TextAlign = cbAlign.SelectedItem.ToString(), Opacity = tbOpacity.Value, 
                    TextRect = new RectangleF(50, 50, 200, 100) 
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
        }

        private void LoadImageToCanvas(string filePath) {
            try {
                Image img = Image.FromFile(filePath);
                canvas = new Bitmap(img.Width, img.Height);
                using (Graphics g = Graphics.FromImage(canvas)) { 
                    g.DrawImage(img, 0, 0); 
                }
                img.Dispose();
                pb.Invalidate();
            } catch { }
        }

        private Rectangle GetDisplayRect() {
            if (canvas == null) return Rectangle.Empty;
            float ratio = Math.Min((float)pb.Width / canvas.Width, (float)pb.Height / canvas.Height);
            int w = (int)(canvas.Width * ratio);
            int h = (int)(canvas.Height * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private void Pb_Paint(object sender, PaintEventArgs e) {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            if (canvas != null) {
                Rectangle disp = GetDisplayRect();
                e.Graphics.DrawImage(canvas, disp);
                
                // 處理座標映射
                Matrix m = new Matrix();
                m.Translate(disp.X, disp.Y);
                m.Scale((float)disp.Width / canvas.Width, (float)disp.Height / canvas.Height);
                e.Graphics.Transform = m;
                
                DrawShapes(e.Graphics);

                e.Graphics.ResetTransform();
            } else {
                StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString("請先載入圖片", this.Font, Brushes.Gray, pb.ClientRectangle, sf);
            }
        }

        private void DrawShapes(Graphics g) {
            foreach (var s in shapes) {
                if (s.Type == "Text") {
                    using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) {
                        g.FillRectangle(bg, s.TextRect);
                    }
                    StringFormat sf = new StringFormat { LineAlignment = StringAlignment.Near };
                    if (s.TextAlign == "置中") sf.Alignment = StringAlignment.Center; 
                    else if (s.TextAlign == "靠右") sf.Alignment = StringAlignment.Far; 
                    else sf.Alignment = StringAlignment.Near;
                    
                    using (SolidBrush tb = new SolidBrush(s.Color)) {
                        g.DrawString(s.Text, s.Font, tb, new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20), sf);
                    }
                } else {
                    using (Pen p = new Pen(s.Color, s.PenWidth)) {
                        int x = Math.Min(s.Start.X, s.End.X);
                        int y = Math.Min(s.Start.Y, s.End.Y);
                        int w = Math.Abs(s.Start.X - s.End.X);
                        int h = Math.Abs(s.Start.Y - s.End.Y);
                        if (s.Type == "Line") g.DrawLine(p, s.Start, s.End);
                        else if (s.Type == "Frame") g.DrawRectangle(p, x, y, w, h);
                        else if (s.Type == "Circle") g.DrawEllipse(p, x, y, w, h);
                    }
                }
            }

            if (drawingShape != null) {
                using (Pen p = new Pen(drawingShape.Color, drawingShape.PenWidth)) {
                    int x = Math.Min(drawingShape.Start.X, drawingShape.End.X);
                    int y = Math.Min(drawingShape.Start.Y, drawingShape.End.Y);
                    int w = Math.Abs(drawingShape.Start.X - drawingShape.End.X);
                    int h = Math.Abs(drawingShape.Start.Y - drawingShape.End.Y);
                    if (drawingShape.Type == "Line") g.DrawLine(p, drawingShape.Start, drawingShape.End);
                    else if (drawingShape.Type == "Frame") g.DrawRectangle(p, x, y, w, h);
                    else if (drawingShape.Type == "Circle") g.DrawEllipse(p, x, y, w, h);
                }
            }
        }

        private Point ScreenToCanvas(Point pt) {
            if (canvas == null) return pt;
            Rectangle disp = GetDisplayRect();
            float ratioX = (float)canvas.Width / disp.Width;
            float ratioY = (float)canvas.Height / disp.Height;
            return new Point((int)((pt.X - disp.X) * ratioX), (int)((pt.Y - disp.Y) * ratioY));
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            CommitTextEdit();
            if (canvas == null) return;
            Point cPt = ScreenToCanvas(e.Location);
            lastMousePos = cPt;

            if (drawMode == "Select") {
                selectedShape = shapes.LastOrDefault(s => s.GetBounds().Contains(cPt));
                if (selectedShape != null) isDraggingShape = true;
            } else {
                drawingShape = new DrawShape { Type = drawMode, Color = penColor, PenWidth = (int)numPenSize.Value, Start = cPt, End = cPt };
            }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            if (canvas == null) return;
            Point cPt = ScreenToCanvas(e.Location);

            if (isDraggingShape && selectedShape != null) {
                selectedShape.Move(cPt.X - lastMousePos.X, cPt.Y - lastMousePos.Y);
                lastMousePos = cPt;
                pb.Invalidate();
            } else if (drawingShape != null) {
                drawingShape.End = cPt;
                pb.Invalidate();
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            isDraggingShape = false;
            if (drawingShape != null) {
                shapes.Add(drawingShape);
                drawingShape = null;
                pb.Invalidate();
            }
        }

        private void Pb_DoubleClick(object sender, EventArgs e) {
            if (canvas == null) return;
            Point cPt = ScreenToCanvas(pb.PointToClient(Cursor.Position));
            var shape = shapes.LastOrDefault(s => s.Type == "Text" && s.GetBounds().Contains(cPt));
            if (shape != null) {
                editingShape = shape;
                Rectangle disp = GetDisplayRect();
                float sX = (float)disp.Width / canvas.Width, sY = (float)disp.Height / canvas.Height;
                RectangleF screenRect = new RectangleF(disp.X + (shape.TextRect.X + 10) * sX, disp.Y + (shape.TextRect.Y + 10) * sY, (shape.TextRect.Width - 20) * sX, (shape.TextRect.Height - 20) * sY);
                
                editBox = new TextBox { 
                    Multiline = true, Text = shape.Text, Font = shape.Font, 
                    Location = Point.Round(screenRect.Location), Size = Size.Round(screenRect.Size) 
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
            if (canvas == null) return;
            shapes.ForEach(s => s.IsSelected = false); 
            Bitmap finalImg = new Bitmap(canvas.Width, canvas.Height);
            using (Graphics g = Graphics.FromImage(finalImg)) { 
                g.DrawImage(canvas, 0, 0); 
                DrawShapes(g); 
            }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "Drawing_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    EncoderParameters ep = new EncoderParameters(1);
                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);
                    ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                    finalImg.Save(sfd.FileName, codec, ep);
                    MessageBox.Show("儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
