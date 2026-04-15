/* * 功能：進階向量繪製模組 (Ctrl+Z 返回, Ctrl+S 儲存)
 * 對應選單名稱：繪製
 * 對應資料表名稱：App_Drawing
 */
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Collections.Generic;

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
                if (Type == "Text") { TextRect.X += dx; TextRect.Y += dy; } 
                else { Start.X += dx; Start.Y += dy; End.X += dx; End.Y += dy; }
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
        
        private Color penColor = Color.Red;
        private string drawMode = "Line";
        private int penWidth = 3;
        
        private bool isDraggingShape = false;
        private bool isResizingText = false;
        private Point lastMousePos;

        private bool isTextModeActive = false;
        private string textContent = "請輸入文字";
        private Font textFont = new Font("Microsoft JhengHei UI", 24, FontStyle.Bold);
        private Color textColor = Color.White, textBgColor = Color.Black, textBorderColor = Color.White;
        private int textOpacity = 150;
        private TextBox txtInput;

        public App_Drawing() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
        }

        private void InitializeUI() {
            Panel ctrlPanel = new Panel { Dock = DockStyle.Top, Height = 130, BackColor = SystemColors.Control };
            
            Button btnLoad = new Button { Text = "載入圖片", Left = 15, Top = 15, Width = 90, Height = 32 };
            Button btnRotate = new Button { Text = "旋轉圖片", Left = 110, Top = 15, Width = 90, Height = 32 };
            Button btnPenColor = new Button { Text = "", Left = 205, Top = 15, Width = 50, Height = 32, BackColor = penColor }; 
            
            ComboBox cbMode = new ComboBox { Left = 265, Top = 16, Width = 90, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font(this.Font.FontFamily, 12) };
            cbMode.Items.AddRange(new string[] { "畫線", "畫框", "畫圓" });
            cbMode.SelectedIndex = 0;

            ComboBox cbSize = new ComboBox { Left = 365, Top = 16, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font(this.Font.FontFamily, 12) };
            cbSize.Items.AddRange(new string[] { "細(2pt)", "中(5pt)", "粗(10pt)" });
            cbSize.SelectedIndex = 0;

            Button btnUndo = new Button { Text = "返回", Left = 455, Top = 15, Width = 70, Height = 32 };
            Button btnClear = new Button { Text = "清除全部", Left = 535, Top = 15, Width = 90, Height = 32, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存圖片", Left = 635, Top = 15, Width = 90, Height = 32, BackColor = Color.SeaGreen, ForeColor = Color.White };

            Button btnInsertText = new Button { Text = "插入文字框", Left = 15, Top = 75, Width = 110, Height = 32, BackColor = Color.SteelBlue, ForeColor = Color.White };
            txtInput = new TextBox { Left = 135, Top = 80, Width = 150, Text = textContent };
            
            Button btnFont = new Button { Text = "字體", Left = 295, Top = 75, Width = 60, Height = 32 };
            Button btnTextColor = new Button { Text = "字色", Left = 360, Top = 75, Width = 60, Height = 32, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Left = 425, Top = 75, Width = 60, Height = 32, BackColor = textBgColor };
            Button btnBorderColor = new Button { Text = "框色", Left = 490, Top = 75, Width = 60, Height = 32, BackColor = textBorderColor };
            
            Label lblOpacity = new Label { Text = "透明度:", Left = 560, Top = 83, AutoSize = true };
            TrackBar tbOpacity = new TrackBar { Left = 620, Top = 75, Width = 120, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };

            btnLoad.Click += (s, e) => LoadImage();
            btnRotate.Click += (s, e) => RotateCanvas();
            btnPenColor.Click += (s, e) => ChooseColor(ref penColor, btnPenColor);
            cbMode.SelectedIndexChanged += (s, e) => { drawMode = cbMode.SelectedIndex == 0 ? "Line" : (cbMode.SelectedIndex == 1 ? "Frame" : "Circle"); };
            cbSize.SelectedIndexChanged += (s, e) => penWidth = cbSize.SelectedIndex == 0 ? 2 : (cbSize.SelectedIndex == 1 ? 5 : 10);
            
            btnUndo.Click += (s, e) => UndoShape();
            btnClear.Click += (s, e) => { shapes.Clear(); if (canvas != null) { canvas.Dispose(); canvas = null; } pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();

            btnInsertText.Click += (s, e) => { isTextModeActive = true; pb.Cursor = Cursors.IBeam; };
            txtInput.TextChanged += (s, e) => { textContent = txtInput.Text; UpdateSelectedTextProperty(); };
            btnFont.Click += (s, e) => { using (FontDialog fd = new FontDialog { Font = textFont }) { if (fd.ShowDialog() == DialogResult.OK) { textFont = fd.Font; UpdateSelectedTextProperty(); } } };
            btnTextColor.Click += (s, e) => { ChooseColor(ref textColor, btnTextColor); UpdateSelectedTextProperty(); };
            btnBgColor.Click += (s, e) => { ChooseColor(ref textBgColor, btnBgColor); UpdateSelectedTextProperty(); };
            btnBorderColor.Click += (s, e) => { ChooseColor(ref textBorderColor, btnBorderColor); UpdateSelectedTextProperty(); };
            tbOpacity.ValueChanged += (s, e) => { textOpacity = tbOpacity.Value; UpdateSelectedTextProperty(); };

            ctrlPanel.Controls.AddRange(new Control[] { 
                btnLoad, btnRotate, btnPenColor, cbMode, cbSize, btnUndo, btnClear, btnSave,
                btnInsertText, txtInput, btnFont, btnTextColor, btnBgColor, btnBorderColor, lblOpacity, tbOpacity
            });

            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.DarkGray };
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

        private void UndoShape() {
            if (shapes.Count > 0) { shapes.RemoveAt(shapes.Count - 1); pb.Invalidate(); }
        }

        private void UpdateSelectedTextProperty() {
            if (selectedShape != null && selectedShape.Type == "Text") {
                selectedShape.Text = textContent; selectedShape.Font = textFont;
                selectedShape.Color = textColor; selectedShape.BgColor = textBgColor;
                selectedShape.BorderColor = textBorderColor; selectedShape.Opacity = textOpacity;
                pb.Invalidate();
            }
        }

        private void LoadImage() {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Image|*.jpg;*.png;*.bmp" }) {
                if (ofd.ShowDialog() == DialogResult.OK) { canvas = new Bitmap(ofd.FileName); shapes.Clear(); pb.Invalidate(); }
            }
        }

        private void RotateCanvas() {
            if (canvas != null) {
                int w = canvas.Width, h = canvas.Height;
                canvas.RotateFlip(RotateFlipType.Rotate90FlipNone);
                foreach (var s in shapes) {
                    if (s.Type == "Text") {
                        float ox = s.TextRect.X, oy = s.TextRect.Y;
                        s.TextRect.X = h - oy - s.TextRect.Height; s.TextRect.Y = ox;
                    } else {
                        int ox1 = s.Start.X, oy1 = s.Start.Y; s.Start.X = h - oy1; s.Start.Y = ox1;
                        int ox2 = s.End.X, oy2 = s.End.Y; s.End.X = h - oy2; s.End.Y = ox2;
                    }
                }
                pb.Invalidate();
            }
        }

        // --- 修正快捷鍵綁定：Ctrl+Z 返回, Ctrl+S 儲存 ---
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            if (keyData == Keys.Delete && !txtInput.Focused) {
                shapes.RemoveAll(s => s.IsSelected); selectedShape = null; pb.Invalidate();
                return true;
            }
            if (keyData == (Keys.Control | Keys.Z)) {
                UndoShape(); return true;
            }
            if (keyData == (Keys.Control | Keys.S)) {
                SaveImage(); return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private Rectangle GetDisplayRect() {
            if (canvas == null) return Rectangle.Empty;
            float ratio = Math.Min((float)pb.Width / canvas.Width, (float)pb.Height / canvas.Height);
            int w = (int)(canvas.Width * ratio), h = (int)(canvas.Height * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private Point TranslatePoint(Point p) {
            Rectangle disp = GetDisplayRect();
            if (disp.IsEmpty) return p;
            float scaleX = (float)canvas.Width / disp.Width, scaleY = (float)canvas.Height / disp.Height;
            return new Point((int)((p.X - disp.X) * scaleX), (int)((p.Y - disp.Y) * scaleY));
        }

        private void DrawShapes(Graphics g) {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            foreach (var s in shapes) {
                if (s.Type == "Text") {
                    using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) g.FillRectangle(bg, s.TextRect);
                    using (Pen border = new Pen(s.BorderColor, 3)) g.DrawRectangle(border, s.TextRect.X, s.TextRect.Y, s.TextRect.Width, s.TextRect.Height);
                    RectangleF textContentRect = new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20);
                    using (SolidBrush tb = new SolidBrush(s.Color)) g.DrawString(s.Text, s.Font, tb, textContentRect);
                } else {
                    using (Pen p = new Pen(s.Color, s.PenWidth)) {
                        int x = Math.Min(s.Start.X, s.End.X), y = Math.Min(s.Start.Y, s.End.Y);
                        int w = Math.Abs(s.Start.X - s.End.X), h = Math.Abs(s.Start.Y - s.End.Y);
                        if (s.Type == "Line") g.DrawLine(p, s.Start, s.End);
                        else if (s.Type == "Frame") g.DrawRectangle(p, x, y, w, h);
                        else if (s.Type == "Circle") g.DrawEllipse(p, x, y, w, h);
                    }
                }
                if (s.IsSelected) {
                    using (Pen dash = new Pen(Color.Cyan, 2) { DashStyle = DashStyle.Dash }) {
                        Rectangle bounds = s.GetBounds(); bounds.Inflate(5, 5); g.DrawRectangle(dash, bounds);
                    }
                    if (s.Type == "Text") {
                        RectangleF handle = s.GetResizeHandle();
                        g.FillRectangle(new SolidBrush(Color.FromArgb(150, Color.Cyan)), handle);
                        g.DrawRectangle(Pens.DarkBlue, handle.X, handle.Y, handle.Width, handle.Height);
                    }
                }
            }
        }

        private void Pb_Paint(object sender, PaintEventArgs e) {
            Rectangle disp = GetDisplayRect();
            if (canvas != null) {
                e.Graphics.DrawImage(canvas, disp);
                float sX = (float)disp.Width / canvas.Width, sY = (float)disp.Height / canvas.Height;
                Matrix m = new Matrix(); m.Translate(disp.X, disp.Y); m.Scale(sX, sY);
                e.Graphics.Transform = m; DrawShapes(e.Graphics); e.Graphics.ResetTransform();
            }
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            pb.Focus();
            if (canvas == null) return;
            Point imgPt = TranslatePoint(e.Location);

            if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) {
                isResizingText = true; lastMousePos = imgPt; return;
            }

            selectedShape = null;
            for (int i = shapes.Count - 1; i >= 0; i--) {
                if (shapes[i].GetBounds().Contains(imgPt)) { selectedShape = shapes[i]; break; }
            }
            foreach (var s in shapes) s.IsSelected = (s == selectedShape);

            if (selectedShape != null) {
                isDraggingShape = true; lastMousePos = imgPt;
                if (selectedShape.Type == "Text") {
                    txtInput.Text = selectedShape.Text; textContent = selectedShape.Text;
                    textFont = selectedShape.Font; textColor = selectedShape.Color;
                    textBgColor = selectedShape.BgColor; textBorderColor = selectedShape.BorderColor;
                    textOpacity = selectedShape.Opacity;
                }
                pb.Invalidate(); return;
            }

            if (isTextModeActive) {
                var s = new DrawShape { Type = "Text", Text = textContent, Font = textFont, Color = textColor, BgColor = textBgColor, BorderColor = textBorderColor, Opacity = textOpacity, TextRect = new RectangleF(imgPt.X, imgPt.Y, 200, 80), IsSelected = true };
                shapes.ForEach(x => x.IsSelected = false);
                shapes.Add(s); selectedShape = s;
                isTextModeActive = false; pb.Cursor = Cursors.Cross; pb.Invalidate(); return;
            }

            drawingShape = new DrawShape { Type = drawMode, Color = penColor, PenWidth = penWidth, Start = imgPt, End = imgPt };
            shapes.Add(drawingShape);
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            Point imgPt = TranslatePoint(e.Location);
            
            if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) {
                pb.Cursor = Cursors.SizeNWSE;
            } else if (!isTextModeActive) {
                pb.Cursor = Cursors.Cross;
            }

            if (isResizingText && selectedShape != null) {
                selectedShape.TextRect.Width = Math.Max(50, imgPt.X - selectedShape.TextRect.X);
                selectedShape.TextRect.Height = Math.Max(30, imgPt.Y - selectedShape.TextRect.Y);
                pb.Invalidate();
            } else if (isDraggingShape && selectedShape != null) {
                selectedShape.Move(imgPt.X - lastMousePos.X, imgPt.Y - lastMousePos.Y);
                lastMousePos = imgPt; pb.Invalidate();
            } else if (drawingShape != null) {
                drawingShape.End = imgPt; pb.Invalidate();
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            isDraggingShape = false; isResizingText = false; drawingShape = null;
        }

        private void SaveImage() {
            if (canvas == null) return;
            shapes.ForEach(s => s.IsSelected = false); 
            Bitmap finalImg = new Bitmap(canvas.Width, canvas.Height);
            using (Graphics g = Graphics.FromImage(finalImg)) { g.DrawImage(canvas, 0, 0); DrawShapes(g); }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg" }) {
                if (sfd.ShowDialog() == DialogResult.OK) { finalImg.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg); MessageBox.Show("儲存成功！"); }
            }
        }
    }
}
