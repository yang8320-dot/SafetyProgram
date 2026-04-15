/* * 功能：拼貼與繪製終極版 (Ctrl+Z 返回, Ctrl+S 儲存)
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

        private int activeFrameIndex = -1;
        private int spacing = 10, baseCanvasSize = 1500; 

        private float div1 = 0.5f, div2 = 0.5f;
        private int activeDivider = 0; 

        private ComboBox cbLayout;
        private TrackBar tbSpacing, tbScale, tbRotate;
        
        private string drawMode = "Line";
        private int penWidth = 3;
        private Color penColor = Color.Red;
        
        private bool isTextModeActive = false;
        private string textContent = "拖曳文字";
        private Font textFont = new Font("Microsoft JhengHei UI", 36, FontStyle.Bold);
        private Color textColor = Color.White, textBgColor = Color.Black, textBorderColor = Color.White;
        private int textOpacity = 150;
        private TextBox txtInput;
        
        private bool isDraggingShape = false, isResizingText = false, isPanningImage = false;
        private Point lastMousePos;

        public App_Collage() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
            LoadTemplate(0); 
        }

        private void InitializeUI() {
            Panel ctrlPanel = new Panel { Dock = DockStyle.Top, Height = 180, BackColor = SystemColors.Control };
            
            Label lblLayout = new Label { Text = "模版:", Left = 15, Top = 20, AutoSize = true };
            cbLayout = new ComboBox { Left = 65, Top = 15, Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
            cbLayout.Items.AddRange(new string[] { "上下兩張", "左右兩張", "上1 下2", "左1 右2", "左2 右1" });
            cbLayout.SelectedIndex = 0;

            Label lblSpacing = new Label { Text = "間距:", Left = 185, Top = 20, AutoSize = true };
            tbSpacing = new TrackBar { Left = 230, Top = 15, Width = 100, Minimum = 0, Maximum = 100, Value = spacing, TickStyle = TickStyle.None };

            Button btnClearFrame = new Button { Text = "清除所選圖片", Left = 340, Top = 12, Width = 110, Height = 32 };
            Button btnClearAll = new Button { Text = "全部清除", Left = 470, Top = 12, Width = 100, Height = 32, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存拼貼圖", Left = 590, Top = 12, Width = 110, Height = 32, BackColor = Color.SeaGreen, ForeColor = Color.White };

            Label lblScale = new Label { Text = "縮放圖片:", Left = 15, Top = 60, AutoSize = true, ForeColor = Color.Blue };
            tbScale = new TrackBar { Left = 90, Top = 55, Width = 150, Minimum = 10, Maximum = 300, Value = 100, TickStyle = TickStyle.None, Enabled = false };
            
            Label lblRotate = new Label { Text = "旋轉圖片:", Left = 250, Top = 60, AutoSize = true, ForeColor = Color.Blue };
            tbRotate = new TrackBar { Left = 330, Top = 55, Width = 150, Minimum = -180, Maximum = 180, Value = 0, TickStyle = TickStyle.None, Enabled = false };

            Button btnInsertText = new Button { Text = "插入文字框", Left = 15, Top = 95, Width = 110, Height = 32, BackColor = Color.SteelBlue, ForeColor = Color.White };
            txtInput = new TextBox { Left = 145, Top = 100, Width = 150, Text = textContent };
            
            Button btnFont = new Button { Text = "字體", Left = 315, Top = 95, Width = 55, Height = 32 };
            Button btnTextColor = new Button { Text = "字色", Left = 380, Top = 95, Width = 55, Height = 32, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Left = 445, Top = 95, Width = 55, Height = 32, BackColor = textBgColor };
            Label lblOpacity = new Label { Text = "透明度:", Left = 510, Top = 105, AutoSize = true };
            TrackBar tbOpacity = new TrackBar { Left = 570, Top = 95, Width = 130, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };

            Label lblDraw = new Label { Text = "繪圖工具:", Left = 15, Top = 145, AutoSize = true, ForeColor = Color.DarkMagenta };
            ComboBox cbMode = new ComboBox { Left = 90, Top = 140, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbMode.Items.AddRange(new string[] { "畫線", "畫框", "畫圓", "關閉繪圖" });
            cbMode.SelectedIndex = 3; 

            ComboBox cbSize = new ComboBox { Left = 180, Top = 140, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbSize.Items.AddRange(new string[] { "細(2pt)", "中(5pt)", "粗(10pt)" });
            cbSize.SelectedIndex = 0;
            
            Button btnPenColor = new Button { Text = "", Left = 270, Top = 138, Width = 40, Height = 30, BackColor = penColor };
            Button btnUndo = new Button { Text = "返回", Left = 320, Top = 138, Width = 70, Height = 30 };

            cbLayout.SelectedIndexChanged += (s, e) => LoadTemplate(cbLayout.SelectedIndex);
            tbSpacing.ValueChanged += (s, e) => { spacing = tbSpacing.Value; pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();

            tbScale.ValueChanged += (s, e) => { if (activeFrameIndex >= 0) { frames[activeFrameIndex].Scale = tbScale.Value / 100f; pb.Invalidate(); } };
            tbRotate.ValueChanged += (s, e) => { if (activeFrameIndex >= 0) { frames[activeFrameIndex].Angle = tbRotate.Value; pb.Invalidate(); } };
            
            btnClearFrame.Click += (s, e) => { if (activeFrameIndex >= 0 && frames[activeFrameIndex].Img != null) { frames[activeFrameIndex].Img.Dispose(); frames[activeFrameIndex].Img = null; pb.Invalidate(); } };
            btnClearAll.Click += (s, e) => { foreach (var f in frames) if (f.Img != null) f.Img.Dispose(); shapes.Clear(); LoadTemplate(cbLayout.SelectedIndex); pb.Invalidate(); };

            btnInsertText.Click += (s, e) => { isTextModeActive = true; pb.Cursor = Cursors.IBeam; };
            txtInput.TextChanged += (s, e) => { textContent = txtInput.Text; UpdateSelectedTextProperty(); };
            btnFont.Click += (s, e) => { using (FontDialog fd = new FontDialog { Font = textFont }) { if (fd.ShowDialog() == DialogResult.OK) { textFont = fd.Font; UpdateSelectedTextProperty(); } } };
            btnTextColor.Click += (s, e) => ChooseColor(ref textColor, btnTextColor);
            btnBgColor.Click += (s, e) => ChooseColor(ref textBgColor, btnBgColor);
            tbOpacity.ValueChanged += (s, e) => { textOpacity = tbOpacity.Value; UpdateSelectedTextProperty(); };

            cbMode.SelectedIndexChanged += (s, e) => { drawMode = cbMode.SelectedIndex == 0 ? "Line" : (cbMode.SelectedIndex == 1 ? "Frame" : (cbMode.SelectedIndex == 2 ? "Circle" : "None")); };
            cbSize.SelectedIndexChanged += (s, e) => penWidth = cbSize.SelectedIndex == 0 ? 2 : (cbSize.SelectedIndex == 1 ? 5 : 10);
            btnPenColor.Click += (s, e) => ChooseColor(ref penColor, btnPenColor);
            btnUndo.Click += (s, e) => UndoShape();

            ctrlPanel.Controls.AddRange(new Control[] { 
                lblLayout, cbLayout, lblSpacing, tbSpacing, btnClearFrame, btnClearAll, btnSave,
                lblScale, tbScale, lblRotate, tbRotate,
                btnInsertText, txtInput, btnFont, btnTextColor, btnBgColor, lblOpacity, tbOpacity,
                lblDraw, cbMode, cbSize, btnPenColor, btnUndo
            });

            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AllowDrop = true };
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
            using (ColorDialog cd = new ColorDialog { Color = target }) { if (cd.ShowDialog() == DialogResult.OK) { target = cd.Color; btn.BackColor = target; pb.Invalidate(); } }
        }

        private void UndoShape() { if (shapes.Count > 0) { shapes.RemoveAt(shapes.Count - 1); pb.Invalidate(); } }
        
        private void UpdateSelectedTextProperty() {
            if (selectedShape != null && selectedShape.Type == "Text") {
                selectedShape.Text = textContent; selectedShape.Font = textFont;
                selectedShape.Color = textColor; selectedShape.BgColor = textBgColor;
                selectedShape.Opacity = textOpacity; pb.Invalidate();
            }
        }

        // --- 修正快捷鍵綁定：Ctrl+Z 返回, Ctrl+S 儲存 ---
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            if (keyData == Keys.Delete && !txtInput.Focused) {
                shapes.RemoveAll(s => s.IsSelected); selectedShape = null; pb.Invalidate(); return true;
            }
            if (keyData == (Keys.Control | Keys.Z)) {
                UndoShape(); return true;
            }
            if (keyData == (Keys.Control | Keys.S)) {
                SaveImage(); return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void LoadTemplate(int index) {
            List<Image> oldImages = new List<Image>();
            foreach (var f in frames) if (f.Img != null) oldImages.Add(f.Img);

            frames.Clear(); activeFrameIndex = -1;
            div1 = 0.5f; div2 = 0.5f; 
            UpdateFramesLayout();

            for (int i = 0; i < Math.Min(oldImages.Count, frames.Count); i++) {
                frames[i].Img = oldImages[i];
                if(frames[i].Img != null) frames[i].Scale = Math.Max(GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize).Width / frames[i].Img.Width, GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize).Height / frames[i].Img.Height);
            }
            UpdateActiveFrameUI(); pb.Invalidate();
        }

        private void UpdateFramesLayout() {
            int mode = cbLayout.SelectedIndex;
            while(frames.Count < 3) frames.Add(new CollageFrame());
            
            if (mode == 0) { 
                frames[0].NormalizedRect = new RectangleF(0, 0, 1, div1);
                frames[1].NormalizedRect = new RectangleF(0, div1, 1, 1 - div1);
                frames.RemoveAt(2);
            } else if (mode == 1) { 
                frames[0].NormalizedRect = new RectangleF(0, 0, div1, 1);
                frames[1].NormalizedRect = new RectangleF(div1, 0, 1 - div1, 1);
                frames.RemoveAt(2);
            } else if (mode == 2) { 
                frames[0].NormalizedRect = new RectangleF(0, 0, 1, div1);
                frames[1].NormalizedRect = new RectangleF(0, div1, div2, 1 - div1);
                frames[2].NormalizedRect = new RectangleF(div2, div1, 1 - div2, 1 - div1);
            } else if (mode == 3) { 
                frames[0].NormalizedRect = new RectangleF(0, 0, div1, 1);
                frames[1].NormalizedRect = new RectangleF(div1, 0, 1 - div1, div2);
                frames[2].NormalizedRect = new RectangleF(div1, div2, 1 - div1, 1 - div2);
            } else if (mode == 4) { 
                frames[0].NormalizedRect = new RectangleF(0, 0, div1, div2);
                frames[1].NormalizedRect = new RectangleF(0, div2, div1, 1 - div2);
                frames[2].NormalizedRect = new RectangleF(div1, 0, 1 - div1, 1);
            }
        }

        private int GetDividerAt(float nx, float ny) {
            float t = 0.02f; 
            int mode = cbLayout.SelectedIndex;
            if (mode == 0 && Math.Abs(ny - div1) < t) return 1;
            if (mode == 1 && Math.Abs(nx - div1) < t) return 1;
            if (mode == 2) { if (Math.Abs(ny - div1) < t) return 1; if (ny > div1 && Math.Abs(nx - div2) < t) return 2; }
            if (mode == 3) { if (Math.Abs(nx - div1) < t) return 1; if (nx > div1 && Math.Abs(ny - div2) < t) return 2; }
            if (mode == 4) { if (Math.Abs(nx - div1) < t) return 1; if (nx < div1 && Math.Abs(ny - div2) < t) return 2; }
            return 0;
        }

        private Rectangle GetDisplayRect() {
            float ratio = Math.Min((float)pb.Width / baseCanvasSize, (float)pb.Height / baseCanvasSize);
            int w = (int)(baseCanvasSize * ratio), h = (int)(baseCanvasSize * ratio);
            return new Rectangle((pb.Width - w) / 2, (pb.Height - h) / 2, w, h);
        }

        private RectangleF GetActualFrameRect(CollageFrame frame, int W, int H) {
            float x = frame.NormalizedRect.X * W + spacing, y = frame.NormalizedRect.Y * H + spacing;
            float w = frame.NormalizedRect.Width * W - (spacing * 2), h = frame.NormalizedRect.Height * H - (spacing * 2);
            return new RectangleF(x, y, w, h);
        }

        private void Pb_Paint(object sender, PaintEventArgs e) {
            Rectangle disp = GetDisplayRect();
            e.Graphics.FillRectangle(Brushes.White, disp); 

            float sX = (float)disp.Width / baseCanvasSize, sY = (float)disp.Height / baseCanvasSize;

            for (int i = 0; i < frames.Count; i++) {
                RectangleF highResRect = GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize);
                RectangleF drawRect = new RectangleF(disp.X + highResRect.X * sX, disp.Y + highResRect.Y * sY, highResRect.Width * sX, highResRect.Height * sY);

                e.Graphics.SetClip(drawRect);
                if (frames[i].Img != null) {
                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    Matrix m = new Matrix();
                    PointF center = new PointF(drawRect.X + drawRect.Width / 2 + (frames[i].OffsetX * sX), drawRect.Y + drawRect.Height / 2 + (frames[i].OffsetY * sY));
                    m.Translate(center.X, center.Y); m.Rotate(frames[i].Angle); m.Scale(frames[i].Scale * sX, frames[i].Scale * sY); m.Translate(-frames[i].Img.Width / 2f, -frames[i].Img.Height / 2f);
                    e.Graphics.Transform = m; e.Graphics.DrawImage(frames[i].Img, Point.Empty); e.Graphics.ResetTransform();
                }
                e.Graphics.ResetClip();
                using (Pen borderPen = new Pen(i == activeFrameIndex ? Color.Red : Color.LightGray, i == activeFrameIndex ? 3 : 1)) { e.Graphics.DrawRectangle(borderPen, drawRect.X, drawRect.Y, drawRect.Width, drawRect.Height); }
            }

            Matrix sm = new Matrix(); sm.Translate(disp.X, disp.Y); sm.Scale(sX, sY);
            e.Graphics.Transform = sm; e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            foreach (var s in shapes) {
                if (s.Type == "Text") {
                    using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) e.Graphics.FillRectangle(bg, s.TextRect);
                    using (Pen border = new Pen(s.BorderColor, 3)) e.Graphics.DrawRectangle(border, s.TextRect.X, s.TextRect.Y, s.TextRect.Width, s.TextRect.Height);
                    using (SolidBrush tb = new SolidBrush(s.Color)) e.Graphics.DrawString(s.Text, s.Font, tb, new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20));
                } else {
                    using (Pen p = new Pen(s.Color, s.PenWidth)) {
                        int x = Math.Min(s.Start.X, s.End.X), y = Math.Min(s.Start.Y, s.End.Y);
                        int w = Math.Abs(s.Start.X - s.End.X), h = Math.Abs(s.Start.Y - s.End.Y);
                        if (s.Type == "Line") e.Graphics.DrawLine(p, s.Start, s.End); else if (s.Type == "Frame") e.Graphics.DrawRectangle(p, x, y, w, h); else if (s.Type == "Circle") e.Graphics.DrawEllipse(p, x, y, w, h);
                    }
                }
                if (s.IsSelected) {
                    using (Pen dash = new Pen(Color.Cyan, 2) { DashStyle = DashStyle.Dash }) { Rectangle b = s.GetBounds(); b.Inflate(5, 5); e.Graphics.DrawRectangle(dash, b); }
                    if (s.Type == "Text") {
                        RectangleF h = s.GetResizeHandle(); e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(150, Color.Cyan)), h); e.Graphics.DrawRectangle(Pens.DarkBlue, h.X, h.Y, h.Width, h.Height);
                    }
                }
            }
            e.Graphics.ResetTransform();
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            pb.Focus();
            Rectangle disp = GetDisplayRect();
            float sX = (float)baseCanvasSize / disp.Width, sY = (float)baseCanvasSize / disp.Height;
            Point imgPt = new Point((int)((e.X - disp.X) * sX), (int)((e.Y - disp.Y) * sY));
            float nx = (float)(e.X - disp.X) / disp.Width, ny = (float)(e.Y - disp.Y) / disp.Height;

            activeDivider = GetDividerAt(nx, ny);
            if (activeDivider > 0) return;

            if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) { isResizingText = true; lastMousePos = imgPt; return; }

            selectedShape = null;
            for (int i = shapes.Count - 1; i >= 0; i--) { if (shapes[i].GetBounds().Contains(imgPt)) { selectedShape = shapes[i]; break; } }
            shapes.ForEach(s => s.IsSelected = (s == selectedShape));

            if (selectedShape != null) {
                isDraggingShape = true; lastMousePos = imgPt;
                if (selectedShape.Type == "Text") { txtInput.Text = selectedShape.Text; textContent = selectedShape.Text; }
                pb.Invalidate(); return;
            }

            if (isTextModeActive) {
                var s = new App_Drawing.DrawShape { Type = "Text", Text = textContent, Font = textFont, Color = textColor, BgColor = textBgColor, BorderColor = textBorderColor, Opacity = textOpacity, TextRect = new RectangleF(imgPt.X, imgPt.Y, 300, 100), IsSelected = true };
                shapes.ForEach(x => x.IsSelected = false); shapes.Add(s); selectedShape = s; isTextModeActive = false; pb.Invalidate(); return;
            }
            if (drawMode != "None") {
                drawingShape = new App_Drawing.DrawShape { Type = drawMode, Color = penColor, PenWidth = penWidth, Start = imgPt, End = imgPt };
                shapes.Add(drawingShape); return;
            }

            int clickedIndex = -1;
            for (int i = 0; i < frames.Count; i++) { if (GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize).Contains(imgPt)) { clickedIndex = i; break; } }
            if (clickedIndex >= 0) {
                activeFrameIndex = clickedIndex; UpdateActiveFrameUI(); pb.Invalidate();
                if (frames[clickedIndex].Img == null) {
                    using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Images|*.jpg;*.png;*.bmp" }) {
                        if (ofd.ShowDialog() == DialogResult.OK) { frames[clickedIndex].Img = Image.FromFile(ofd.FileName); frames[clickedIndex].Scale = Math.Max(GetActualFrameRect(frames[clickedIndex], baseCanvasSize, baseCanvasSize).Width / frames[clickedIndex].Img.Width, GetActualFrameRect(frames[clickedIndex], baseCanvasSize, baseCanvasSize).Height / frames[clickedIndex].Img.Height); pb.Invalidate(); }
                    }
                } else { isPanningImage = true; lastMousePos = e.Location; }
            } else { activeFrameIndex = -1; UpdateActiveFrameUI(); pb.Invalidate(); }
        }

        private void UpdateActiveFrameUI() {
            if (activeFrameIndex >= 0) { tbScale.Enabled = true; tbRotate.Enabled = true; tbScale.Value = (int)(frames[activeFrameIndex].Scale * 100); tbRotate.Value = (int)frames[activeFrameIndex].Angle; } 
            else { tbScale.Enabled = false; tbRotate.Enabled = false; tbScale.Value = 100; tbRotate.Value = 0; }
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            Rectangle disp = GetDisplayRect();
            float nx = (float)(e.X - disp.X) / disp.Width, ny = (float)(e.Y - disp.Y) / disp.Height;
            float sX = (float)baseCanvasSize / disp.Width, sY = (float)baseCanvasSize / disp.Height;
            Point imgPt = new Point((int)((e.X - disp.X) * sX), (int)((e.Y - disp.Y) * sY));

            int hoverDiv = GetDividerAt(nx, ny);
            if (hoverDiv > 0 || activeDivider > 0) pb.Cursor = (cbLayout.SelectedIndex == 1 || (cbLayout.SelectedIndex > 1 && hoverDiv == 1)) ? Cursors.VSplit : Cursors.HSplit;
            else if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) pb.Cursor = Cursors.SizeNWSE;
            else pb.Cursor = drawMode != "None" ? Cursors.Cross : Cursors.Default;

            if (activeDivider > 0) {
                if (activeDivider == 1) div1 = Math.Max(0.1f, Math.Min(0.9f, (cbLayout.SelectedIndex == 1 || cbLayout.SelectedIndex > 2) ? nx : ny));
                else div2 = Math.Max(0.1f, Math.Min(0.9f, (cbLayout.SelectedIndex == 2) ? nx : ny));
                UpdateFramesLayout(); pb.Invalidate();
            } else if (isResizingText && selectedShape != null) {
                selectedShape.TextRect.Width = Math.Max(100, imgPt.X - selectedShape.TextRect.X); selectedShape.TextRect.Height = Math.Max(50, imgPt.Y - selectedShape.TextRect.Y); pb.Invalidate();
            } else if (isDraggingShape && selectedShape != null) {
                selectedShape.Move(imgPt.X - lastMousePos.X, imgPt.Y - lastMousePos.Y); lastMousePos = imgPt; pb.Invalidate();
            } else if (drawingShape != null) {
                drawingShape.End = imgPt; pb.Invalidate();
            } else if (isPanningImage && activeFrameIndex >= 0) {
                frames[activeFrameIndex].OffsetX += (e.X - lastMousePos.X) * sX; frames[activeFrameIndex].OffsetY += (e.Y - lastMousePos.Y) * sY; lastMousePos = e.Location; pb.Invalidate();
            }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) { activeDivider = 0; isDraggingShape = false; isResizingText = false; isPanningImage = false; drawingShape = null; }

        private void Pb_DragDrop(object sender, DragEventArgs e) {
            Point pt = pb.PointToClient(new Point(e.X, e.Y));
            Rectangle disp = GetDisplayRect();
            Point imgPt = new Point((int)((pt.X - disp.X) * (baseCanvasSize / disp.Width)), (int)((pt.Y - disp.Y) * (baseCanvasSize / disp.Height)));
            int dropIndex = -1;
            for (int i = 0; i < frames.Count; i++) { if (GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize).Contains(imgPt)) { dropIndex = i; break; } }
            if (dropIndex >= 0) {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].ToLower().EndsWith(".jpg") || files[0].ToLower().EndsWith(".png"))) {
                    if (frames[dropIndex].Img != null) frames[dropIndex].Img.Dispose();
                    frames[dropIndex].Img = Image.FromFile(files[0]);
                    frames[dropIndex].Scale = Math.Max(GetActualFrameRect(frames[dropIndex], baseCanvasSize, baseCanvasSize).Width / frames[dropIndex].Img.Width, GetActualFrameRect(frames[dropIndex], baseCanvasSize, baseCanvasSize).Height / frames[dropIndex].Img.Height);
                    activeFrameIndex = dropIndex; UpdateActiveFrameUI(); pb.Invalidate();
                }
            }
        }

        private void SaveImage() {
            shapes.ForEach(s => s.IsSelected = false); pb.Invalidate();
            Bitmap finalImg = new Bitmap(baseCanvasSize, baseCanvasSize);
            using (Graphics g = Graphics.FromImage(finalImg)) {
                g.FillRectangle(Brushes.White, 0, 0, baseCanvasSize, baseCanvasSize); g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                foreach (var frame in frames) {
                    RectangleF hRect = GetActualFrameRect(frame, baseCanvasSize, baseCanvasSize);
                    g.SetClip(hRect);
                    if (frame.Img != null) {
                        Matrix m = new Matrix(); PointF c = new PointF(hRect.X + hRect.Width / 2 + frame.OffsetX, hRect.Y + hRect.Height / 2 + frame.OffsetY);
                        m.Translate(c.X, c.Y); m.Rotate(frame.Angle); m.Scale(frame.Scale, frame.Scale); m.Translate(-frame.Img.Width / 2f, -frame.Img.Height / 2f);
                        g.Transform = m; g.DrawImage(frame.Img, Point.Empty); g.ResetTransform();
                    }
                    g.ResetClip();
                }
                g.SmoothingMode = SmoothingMode.AntiAlias;
                foreach (var s in shapes) {
                    if (s.Type == "Text") {
                        using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) g.FillRectangle(bg, s.TextRect);
                        using (Pen border = new Pen(s.BorderColor, 3)) g.DrawRectangle(border, s.TextRect.X, s.TextRect.Y, s.TextRect.Width, s.TextRect.Height);
                        using (SolidBrush tb = new SolidBrush(s.Color)) g.DrawString(s.Text, s.Font, tb, new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20));
                    } else {
                        using (Pen p = new Pen(s.Color, s.PenWidth)) {
                            int x = Math.Min(s.Start.X, s.End.X), y = Math.Min(s.Start.Y, s.End.Y);
                            int w = Math.Abs(s.Start.X - s.End.X), h = Math.Abs(s.Start.Y - s.End.Y);
                            if (s.Type == "Line") g.DrawLine(p, s.Start, s.End); else if (s.Type == "Frame") g.DrawRectangle(p, x, y, w, h); else if (s.Type == "Circle") g.DrawEllipse(p, x, y, w, h);
                        }
                    }
                }
            }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "Collage_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    EncoderParameters ep = new EncoderParameters(1); ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L); 
                    ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                    finalImg.Save(sfd.FileName, codec, ep); MessageBox.Show("高解析度拼貼圖儲存成功！");
                }
            }
        }
    }
}
