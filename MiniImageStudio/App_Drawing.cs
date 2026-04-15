/* * 功能：進階向量繪製模組 (雙擊編輯文字、對齊、流式排版、ESC/Delete/Ctrl+S快捷鍵)
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
            FlowLayoutPanel ctrlPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, BackColor = SystemColors.Control };

            // 群組 1: 畫布控制
            GroupBox gb1 = new GroupBox { Text = "畫布", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl1 = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            Button btnLoad = new Button { Text = "載入圖片", Width = 90, Height = 32 };
            Button btnRotate = new Button { Text = "旋轉圖片", Width = 90, Height = 32 };
            Button btnClear = new Button { Text = "清除全部", Width = 90, Height = 32, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存圖片", Width = 90, Height = 32, BackColor = Color.SeaGreen, ForeColor = Color.White };
            fl1.Controls.AddRange(new Control[] { btnLoad, btnRotate, btnClear, btnSave });
            gb1.Controls.Add(fl1);

            // 群組 2: 繪圖工具
            GroupBox gb2 = new GroupBox { Text = "繪圖工具", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl2 = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            cbMode = new ComboBox { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font(this.Font.FontFamily, 12), Margin = new Padding(3,5,3,3) };
            cbMode.Items.AddRange(new string[] { "選取", "畫框", "畫線", "畫圓" });
            cbMode.SelectedIndex = 1; // 預設畫框
            
            Label lblSize = new Label { Text = "粗細:", AutoSize = true, Margin = new Padding(3,10,3,3) };
            numPenSize = new NumericUpDown { Minimum = 1, Maximum = 10, Value = 5, Width = 50, Font = new Font(this.Font.FontFamily, 12), Margin = new Padding(3,5,3,3) };
            Button btnPenColor = new Button { Width = 32, Height = 32, BackColor = penColor }; 
            Button btnUndo = new Button { Text = "返回", Width = 70, Height = 32 };
            fl2.Controls.AddRange(new Control[] { cbMode, lblSize, numPenSize, btnPenColor, btnUndo });
            gb2.Controls.Add(fl2);

            // 群組 3: 文字工具
            GroupBox gb3 = new GroupBox { Text = "文字工具 (雙擊文字框可編輯)", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl3 = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            Button btnInsertText = new Button { Text = "插入文字框", Width = 110, Height = 32, BackColor = Color.SteelBlue, ForeColor = Color.White };
            
            cbAlign = new ComboBox { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3,5,3,3) };
            cbAlign.Items.AddRange(new string[] { "靠左", "置中", "靠右" });
            cbAlign.SelectedIndex = 0;

            Button btnFont = new Button { Text = "字體", Width = 60, Height = 32 };
            Button btnTextColor = new Button { Text = "字色", Width = 60, Height = 32, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Width = 60, Height = 32, BackColor = textBgColor };
            Button btnBorderColor = new Button { Text = "框色", Width = 60, Height = 32, BackColor = textBorderColor };
            Label lblOpacity = new Label { Text = "透明度:", AutoSize = true, Margin = new Padding(3,10,3,3) };
            TrackBar tbOpacity = new TrackBar { Width = 100, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };
            fl3.Controls.AddRange(new Control[] { btnInsertText, cbAlign, btnFont, btnTextColor, btnBgColor, btnBorderColor, lblOpacity, tbOpacity });
            gb3.Controls.Add(fl3);

            ctrlPanel.Controls.AddRange(new Control[] { gb1, gb2, gb3 });

            // 事件綁定
            btnLoad.Click += (s, e) => LoadImage();
            btnRotate.Click += (s, e) => RotateCanvas();
            btnClear.Click += (s, e) => { shapes.Clear(); if (canvas != null) { canvas.Dispose(); canvas = null; } pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();

            cbMode.SelectedIndexChanged += (s, e) => {
                if (cbMode.SelectedIndex == 0) drawMode = "Select";
                else if (cbMode.SelectedIndex == 1) drawMode = "Frame";
                else if (cbMode.SelectedIndex == 2) drawMode = "Line";
                else drawMode = "Circle";
            };
            btnPenColor.Click += (s, e) => ChooseColor(ref penColor, btnPenColor);
            btnUndo.Click += (s, e) => UndoShape();

            btnInsertText.Click += (s, e) => { isTextModeActive = true; cbMode.SelectedIndex = 0; pb.Cursor = Cursors.Cross; };
            cbAlign.SelectedIndexChanged += (s, e) => UpdateSelectedTextProperty();
            btnFont.Click += (s, e) => { using (FontDialog fd = new FontDialog { Font = textFont }) { if (fd.ShowDialog() == DialogResult.OK) { textFont = fd.Font; UpdateSelectedTextProperty(); } } };
            btnTextColor.Click += (s, e) => { ChooseColor(ref textColor, btnTextColor); UpdateSelectedTextProperty(); };
            btnBgColor.Click += (s, e) => { ChooseColor(ref textBgColor, btnBgColor); UpdateSelectedTextProperty(); };
            btnBorderColor.Click += (s, e) => { ChooseColor(ref textBorderColor, btnBorderColor); UpdateSelectedTextProperty(); };
            tbOpacity.ValueChanged += (s, e) => { textOpacity = tbOpacity.Value; UpdateSelectedTextProperty(); };

            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.DarkGray };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;
            pb.MouseDoubleClick += Pb_MouseDoubleClick;

            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(ctrlPanel);
        }

        private void ChooseColor(ref Color target, Button btn) {
            using (ColorDialog cd = new ColorDialog { Color = target }) { if (cd.ShowDialog() == DialogResult.OK) { target = cd.Color; btn.BackColor = target; pb.Invalidate(); } }
        }

        private void UndoShape() { CommitTextEdit(); if (shapes.Count > 0) { shapes.RemoveAt(shapes.Count - 1); pb.Invalidate(); } }

        private void UpdateSelectedTextProperty() {
            if (selectedShape != null && selectedShape.Type == "Text") {
                selectedShape.Font = textFont; selectedShape.Color = textColor;
                selectedShape.BgColor = textBgColor; selectedShape.BorderColor = textBorderColor;
                selectedShape.Opacity = textOpacity; selectedShape.TextAlign = cbAlign.SelectedItem.ToString();
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
                CommitTextEdit();
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

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            if (keyData == Keys.Escape) {
                CommitTextEdit(); drawMode = "Select"; cbMode.SelectedIndex = 0;
                isTextModeActive = false; pb.Cursor = Cursors.Default; pb.Invalidate(); return true;
            }
            if (keyData == Keys.Delete && editBox == null) {
                if (selectedShape != null) { shapes.Remove(selectedShape); selectedShape = null; pb.Invalidate(); return true; }
                if (canvas != null) { canvas.Dispose(); canvas = null; pb.Invalidate(); return true; } // 清除背景
            }
            if (keyData == (Keys.Control | Keys.Z)) { UndoShape(); return true; }
            if (keyData == (Keys.Control | Keys.S)) { SaveImage(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private Rectangle GetDisplayRect() {
            if (canvas == null) return Rectangle.Empty;
            float ratio = Math.Min((float)pb.Width / canvas.Width, (float)pb.Height / canvas.Height);
            int w = (int)(canvas.Width * ratio), h = (int)(canvas.Height * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private Point TranslatePoint(Point p) {
            Rectangle disp = GetDisplayRect(); if (disp.IsEmpty) return p;
            float sX = (float)canvas.Width / disp.Width, sY = (float)canvas.Height / disp.Height;
            return new Point((int)((p.X - disp.X) * sX), (int)((p.Y - disp.Y) * sY));
        }

        private void DrawShapes(Graphics g) {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            foreach (var s in shapes) {
                if (s.Type == "Text") {
                    using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) g.FillRectangle(bg, s.TextRect);
                    using (Pen border = new Pen(s.BorderColor, 3)) g.DrawRectangle(border, s.TextRect.X, s.TextRect.Y, s.TextRect.Width, s.TextRect.Height);
                    
                    StringFormat sf = new StringFormat { LineAlignment = StringAlignment.Near };
                    if (s.TextAlign == "置中") sf.Alignment = StringAlignment.Center;
                    else if (s.TextAlign == "靠右") sf.Alignment = StringAlignment.Far;
                    else sf.Alignment = StringAlignment.Near;

                    RectangleF tRect = new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20);
                    // 隱藏正在編輯的文字
                    if (s != editingShape) {
                        using (SolidBrush tb = new SolidBrush(s.Color)) g.DrawString(s.Text, s.Font, tb, tRect, sf);
                    }
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
                    using (Pen dash = new Pen(Color.Cyan, 2) { DashStyle = DashStyle.Dash }) { Rectangle bounds = s.GetBounds(); bounds.Inflate(5, 5); g.DrawRectangle(dash, bounds); }
                    if (s.Type == "Text") {
                        RectangleF handle = s.GetResizeHandle();
                        g.FillRectangle(new SolidBrush(Color.FromArgb(150, Color.Cyan)), handle); g.DrawRectangle(Pens.DarkBlue, handle.X, handle.Y, handle.Width, handle.Height);
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
            pb.Focus(); CommitTextEdit();
            if (canvas == null) return;
            Point imgPt = TranslatePoint(e.Location);

            if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) {
                isResizingText = true; lastMousePos = imgPt; return;
            }

            selectedShape = null;
            for (int i = shapes.Count - 1; i >= 0; i--) { if (shapes[i].GetBounds().Contains(imgPt)) { selectedShape = shapes[i]; break; } }
            shapes.ForEach(s => s.IsSelected = (s == selectedShape));

            if (selectedShape != null) {
                isDraggingShape = true; lastMousePos = imgPt;
                if (selectedShape.Type == "Text") {
                    textFont = selectedShape.Font; textColor = selectedShape.Color;
                    textBgColor = selectedShape.BgColor; textBorderColor = selectedShape.BorderColor;
                    textOpacity = selectedShape.Opacity; cbAlign.SelectedItem = selectedShape.TextAlign;
                }
                pb.Invalidate(); return;
            }

            if (isTextModeActive) {
                var s = new DrawShape { Type = "Text", Text = textContent, Font = textFont, Color = textColor, BgColor = textBgColor, BorderColor = textBorderColor, Opacity = textOpacity, TextAlign = cbAlign.SelectedItem.ToString(), TextRect = new RectangleF(imgPt.X, imgPt.Y, 200, 80), IsSelected = true };
                shapes.ForEach(x => x.IsSelected = false); shapes.Add(s); selectedShape = s;
                isTextModeActive = false; pb.Cursor = Cursors.Default; pb.Invalidate(); return;
            }
            if (drawMode != "Select") {
                drawingShape = new DrawShape { Type = drawMode, Color = penColor, PenWidth = (int)numPenSize.Value, Start = imgPt, End = imgPt };
                shapes.Add(drawingShape);
            }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            Point imgPt = TranslatePoint(e.Location);
            if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) pb.Cursor = Cursors.SizeNWSE;
            else if (isTextModeActive) pb.Cursor = Cursors.Cross;
            else pb.Cursor = Cursors.Default;

            if (isResizingText && selectedShape != null) {
                selectedShape.TextRect.Width = Math.Max(50, imgPt.X - selectedShape.TextRect.X);
                selectedShape.TextRect.Height = Math.Max(30, imgPt.Y - selectedShape.TextRect.Y);
                pb.Invalidate();
            } else if (isDraggingShape && selectedShape != null) {
                selectedShape.Move(imgPt.X - lastMousePos.X, imgPt.Y - lastMousePos.Y); lastMousePos = imgPt; pb.Invalidate();
            } else if (drawingShape != null) { drawingShape.End = imgPt; pb.Invalidate(); }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) { isDraggingShape = false; isResizingText = false; drawingShape = null; }

        private void Pb_MouseDoubleClick(object sender, MouseEventArgs e) {
            if (selectedShape != null && selectedShape.Type == "Text") ShowEditBox(selectedShape);
        }

        private void ShowEditBox(DrawShape s) {
            editingShape = s;
            editBox = new TextBox { Multiline = true, Text = s.Text, Font = s.Font };
            if (s.TextAlign == "置中") editBox.TextAlign = HorizontalAlignment.Center;
            else if (s.TextAlign == "靠右") editBox.TextAlign = HorizontalAlignment.Right;
            else editBox.TextAlign = HorizontalAlignment.Left;

            Rectangle disp = GetDisplayRect();
            float sX = (float)disp.Width / canvas.Width, sY = (float)disp.Height / canvas.Height;
            RectangleF screenRect = new RectangleF(disp.X + (s.TextRect.X + 10) * sX, disp.Y + (s.TextRect.Y + 10) * sY, (s.TextRect.Width - 20) * sX, (s.TextRect.Height - 20) * sY);
            
            editBox.Location = Point.Round(screenRect.Location); editBox.Size = Size.Round(screenRect.Size);
            editBox.LostFocus += (sender, e) => CommitTextEdit();
            pb.Controls.Add(editBox); editBox.BringToFront(); editBox.Focus(); pb.Invalidate();
        }

        private void CommitTextEdit() {
            if (editBox != null && editingShape != null) {
                editingShape.Text = editBox.Text; pb.Controls.Remove(editBox); editBox.Dispose(); editBox = null; editingShape = null; pb.Invalidate();
            }
        }

        private void SaveImage() {
            CommitTextEdit(); if (canvas == null) return;
            shapes.ForEach(s => s.IsSelected = false); 
            Bitmap finalImg = new Bitmap(canvas.Width, canvas.Height);
            using (Graphics g = Graphics.FromImage(finalImg)) { g.DrawImage(canvas, 0, 0); DrawShapes(g); }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg" }) {
                if (sfd.ShowDialog() == DialogResult.OK) { finalImg.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg); MessageBox.Show("儲存成功！"); }
            }
        }
    }
}
