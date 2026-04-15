/* * 功能：拼貼與繪製終極版 (一大四小框群組化排版、修改預設提示字)
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

        private ComboBox cbLayout, cbAlign, cbMode;
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
            FlowLayoutPanel mainFlow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, BackColor = SystemColors.Control };

            GroupBox gb1 = new GroupBox { Text = "模版與版面", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl1 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            cbLayout = new ComboBox { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3,5,3,3) };
            cbLayout.Items.AddRange(new string[] { "上下兩張", "左右兩張", "上1 下2", "左1 右2", "左2 右1" }); cbLayout.SelectedIndex = 0;
            Label lblSpacing = new Label { Text = "間距:", AutoSize = true, Margin = new Padding(3,8,3,3) };
            tbSpacing = new TrackBar { Width = 100, Minimum = 0, Maximum = 100, Value = spacing, TickStyle = TickStyle.None };
            Button btnClearAll = new Button { Text = "全部清除", Width = 90, Height = 32, BackColor = Color.IndianRed, ForeColor = Color.White };
            Button btnSave = new Button { Text = "儲存拼貼圖", Width = 100, Height = 32, BackColor = Color.SeaGreen, ForeColor = Color.White };
            fl1.Controls.AddRange(new Control[] { cbLayout, lblSpacing, tbSpacing, btnClearAll, btnSave });
            gb1.Controls.Add(fl1);

            GroupBox gb2 = new GroupBox { Text = "選取圖片控制", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl2 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            Label lblScale = new Label { Text = "縮放:", AutoSize = true, ForeColor = Color.Blue, Margin = new Padding(3,8,3,3) };
            tbScale = new TrackBar { Width = 120, Minimum = 10, Maximum = 300, Value = 100, TickStyle = TickStyle.None, Enabled = false };
            Label lblRotate = new Label { Text = "旋轉:", AutoSize = true, ForeColor = Color.Blue, Margin = new Padding(3,8,3,3) };
            tbRotate = new TrackBar { Width = 120, Minimum = -180, Maximum = 180, Value = 0, TickStyle = TickStyle.None, Enabled = false };
            Button btnClearFrame = new Button { Text = "刪除圖片", Width = 80, Height = 32 };
            fl2.Controls.AddRange(new Control[] { lblScale, tbScale, lblRotate, tbRotate, btnClearFrame });
            gb2.Controls.Add(fl2);

            GroupBox gb3 = new GroupBox { Text = "文字工具 (雙擊框可編輯)", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl3 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            Button btnInsertText = new Button { Text = "插入文字框", Width = 100, Height = 32, BackColor = Color.SteelBlue, ForeColor = Color.White };
            cbAlign = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3,5,3,3) };
            cbAlign.Items.AddRange(new string[] { "靠左", "置中", "靠右" }); cbAlign.SelectedIndex = 0;
            Button btnFont = new Button { Text = "字體", Width = 55, Height = 32 };
            Button btnTextColor = new Button { Text = "字色", Width = 55, Height = 32, BackColor = textColor };
            Button btnBgColor = new Button { Text = "底色", Width = 55, Height = 32, BackColor = textBgColor };
            Label lblOpacity = new Label { Text = "透明度:", AutoSize = true, Margin = new Padding(3,8,3,3) };
            TrackBar tbOpacity = new TrackBar { Width = 100, Minimum = 0, Maximum = 255, Value = textOpacity, TickStyle = TickStyle.None };
            fl3.Controls.AddRange(new Control[] { btnInsertText, cbAlign, btnFont, btnTextColor, btnBgColor, lblOpacity, tbOpacity });
            gb3.Controls.Add(fl3);

            GroupBox gb4 = new GroupBox { Text = "繪圖工具", AutoSize = true, Padding = new Padding(5) };
            FlowLayoutPanel fl4 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            cbMode = new ComboBox { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3,5,3,3) };
            cbMode.Items.AddRange(new string[] { "選取", "畫框", "畫線", "畫圓" }); cbMode.SelectedIndex = 0;
            numPenSize = new NumericUpDown { Minimum = 1, Maximum = 10, Value = 5, Width = 50, Margin = new Padding(3,5,3,3) };
            Button btnPenColor = new Button { Width = 32, Height = 32, BackColor = penColor };
            Button btnUndo = new Button { Text = "返回", Width = 70, Height = 32 };
            fl4.Controls.AddRange(new Control[] { cbMode, numPenSize, btnPenColor, btnUndo });
            gb4.Controls.Add(fl4);

            mainFlow.Controls.AddRange(new Control[] { gb1, gb2, gb3, gb4 });

            cbLayout.SelectedIndexChanged += (s, e) => LoadTemplate(cbLayout.SelectedIndex);
            tbSpacing.ValueChanged += (s, e) => { spacing = tbSpacing.Value; pb.Invalidate(); };
            btnSave.Click += (s, e) => SaveImage();
            tbScale.ValueChanged += (s, e) => { if (activeFrameIndex >= 0) { frames[activeFrameIndex].Scale = tbScale.Value / 100f; pb.Invalidate(); } };
            tbRotate.ValueChanged += (s, e) => { if (activeFrameIndex >= 0) { frames[activeFrameIndex].Angle = tbRotate.Value; pb.Invalidate(); } };
            
            btnClearFrame.Click += (s, e) => { if (activeFrameIndex >= 0 && frames[activeFrameIndex].Img != null) { frames[activeFrameIndex].Img.Dispose(); frames[activeFrameIndex].Img = null; pb.Invalidate(); } };
            btnClearAll.Click += (s, e) => { 
                var toDispose = frames.Where(f => f.Img != null).Select(f => f.Img).ToList();
                frames.Clear(); shapes.Clear(); CommitTextEdit(); LoadTemplate(cbLayout.SelectedIndex); pb.Invalidate(); pb.Update();
                toDispose.ForEach(i => i.Dispose());
            };

            btnInsertText.Click += (s, e) => { isTextModeActive = true; cbMode.SelectedIndex = 0; pb.Cursor = Cursors.Cross; };
            cbAlign.SelectedIndexChanged += (s, e) => UpdateSelectedTextProperty();
            btnFont.Click += (s, e) => { using (FontDialog fd = new FontDialog { Font = textFont }) { if (fd.ShowDialog() == DialogResult.OK) { textFont = fd.Font; UpdateSelectedTextProperty(); } } };
            btnTextColor.Click += (s, e) => { ChooseColor(ref textColor, btnTextColor); UpdateSelectedTextProperty(); };
            btnBgColor.Click += (s, e) => { ChooseColor(ref textBgColor, btnBgColor); UpdateSelectedTextProperty(); };
            tbOpacity.ValueChanged += (s, e) => { textOpacity = tbOpacity.Value; UpdateSelectedTextProperty(); };

            cbMode.SelectedIndexChanged += (s, e) => {
                if (cbMode.SelectedIndex == 0) drawMode = "Select";
                else if (cbMode.SelectedIndex == 1) drawMode = "Frame";
                else if (cbMode.SelectedIndex == 2) drawMode = "Line";
                else drawMode = "Circle";
            };
            btnPenColor.Click += (s, e) => ChooseColor(ref penColor, btnPenColor);
            btnUndo.Click += (s, e) => UndoShape();

            pb = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.WhiteSmoke, AllowDrop = true };
            pb.Paint += Pb_Paint;
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;
            pb.MouseDoubleClick += Pb_MouseDoubleClick;
            pb.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            pb.DragDrop += Pb_DragDrop;

            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(mainFlow);
        }

        private void ChooseColor(ref Color target, Button btn) {
            using (ColorDialog cd = new ColorDialog { Color = target }) { if (cd.ShowDialog() == DialogResult.OK) { target = cd.Color; btn.BackColor = target; pb.Invalidate(); } }
        }

        private void UndoShape() { CommitTextEdit(); if (shapes.Count > 0) { shapes.RemoveAt(shapes.Count - 1); pb.Invalidate(); } }
        
        private void UpdateSelectedTextProperty() {
            if (selectedShape != null && selectedShape.Type == "Text") {
                selectedShape.Font = textFont; selectedShape.Color = textColor;
                selectedShape.BgColor = textBgColor; selectedShape.Opacity = textOpacity;
                selectedShape.TextAlign = cbAlign.SelectedItem.ToString(); pb.Invalidate();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            if (keyData == Keys.Escape) {
                CommitTextEdit(); drawMode = "Select"; cbMode.SelectedIndex = 0;
                isTextModeActive = false; pb.Cursor = Cursors.Default; pb.Invalidate(); return true;
            }
            if (keyData == Keys.Delete && editBox == null) {
                if (selectedShape != null) { shapes.Remove(selectedShape); selectedShape = null; pb.Invalidate(); return true; }
                if (activeFrameIndex >= 0 && frames[activeFrameIndex].Img != null) { frames[activeFrameIndex].Img.Dispose(); frames[activeFrameIndex].Img = null; pb.Invalidate(); return true; }
            }
            if (keyData == (Keys.Control | Keys.Z)) { UndoShape(); return true; }
            if (keyData == (Keys.Control | Keys.S)) { SaveImage(); return true; }
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
            
            if (mode == 0) { frames[0].NormalizedRect = new RectangleF(0, 0, 1, div1); frames[1].NormalizedRect = new RectangleF(0, div1, 1, 1 - div1); frames.RemoveAt(2); } 
            else if (mode == 1) { frames[0].NormalizedRect = new RectangleF(0, 0, div1, 1); frames[1].NormalizedRect = new RectangleF(div1, 0, 1 - div1, 1); frames.RemoveAt(2); } 
            else if (mode == 2) { frames[0].NormalizedRect = new RectangleF(0, 0, 1, div1); frames[1].NormalizedRect = new RectangleF(0, div1, div2, 1 - div1); frames[2].NormalizedRect = new RectangleF(div2, div1, 1 - div2, 1 - div1); } 
            else if (mode == 3) { frames[0].NormalizedRect = new RectangleF(0, 0, div1, 1); frames[1].NormalizedRect = new RectangleF(div1, 0, 1 - div1, div2); frames[2].NormalizedRect = new RectangleF(div1, div2, 1 - div1, 1 - div2); } 
            else if (mode == 4) { frames[0].NormalizedRect = new RectangleF(0, 0, div1, div2); frames[1].NormalizedRect = new RectangleF(0, div2, div1, 1 - div2); frames[2].NormalizedRect = new RectangleF(div1, 0, 1 - div1, 1); }
        }

        private int GetDividerAt(float nx, float ny) {
            float t = 0.02f; 
            int mode = cbLayout.SelectedIndex;
            if (mode == 0 && Math.Abs(ny - div1) < t) return 1; if (mode == 1 && Math.Abs(nx - div1) < t) return 1;
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
                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic; Matrix m = new Matrix();
                    PointF center = new PointF(drawRect.X + drawRect.Width / 2 + (frames[i].OffsetX * sX), drawRect.Y + drawRect.Height / 2 + (frames[i].OffsetY * sY));
                    m.Translate(center.X, center.Y); m.Rotate(frames[i].Angle); m.Scale(frames[i].Scale * sX, frames[i].Scale * sY); m.Translate(-frames[i].Img.Width / 2f, -frames[i].Img.Height / 2f);
                    e.Graphics.Transform = m; e.Graphics.DrawImage(frames[i].Img, Point.Empty); e.Graphics.ResetTransform();
                } else {
                    using(StringFormat sf = new StringFormat{Alignment=StringAlignment.Center, LineAlignment=StringAlignment.Center}) {
                        // 更新了您指定的文字
                        e.Graphics.DrawString("請點選插入圖片 或 拖曳上傳圖片", MainForm.UI_Font, Brushes.Gray, drawRect, sf);
                    }
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
                    
                    StringFormat sf = new StringFormat { LineAlignment = StringAlignment.Near };
                    if (s.TextAlign == "置中") sf.Alignment = StringAlignment.Center; else if (s.TextAlign == "靠右") sf.Alignment = StringAlignment.Far; else sf.Alignment = StringAlignment.Near;
                    
                    if (s != editingShape) {
                        using (SolidBrush tb = new SolidBrush(s.Color)) e.Graphics.DrawString(s.Text, s.Font, tb, new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20), sf);
                    }
                } else {
                    using (Pen p = new Pen(s.Color, s.PenWidth)) {
                        int x = Math.Min(s.Start.X, s.End.X), y = Math.Min(s.Start.Y, s.End.Y), w = Math.Abs(s.Start.X - s.End.X), h = Math.Abs(s.Start.Y - s.End.Y);
                        if (s.Type == "Line") e.Graphics.DrawLine(p, s.Start, s.End); else if (s.Type == "Frame") e.Graphics.DrawRectangle(p, x, y, w, h); else if (s.Type == "Circle") e.Graphics.DrawEllipse(p, x, y, w, h);
                    }
                }
                if (s.IsSelected) {
                    using (Pen dash = new Pen(Color.Cyan, 2) { DashStyle = DashStyle.Dash }) { Rectangle b = s.GetBounds(); b.Inflate(5, 5); e.Graphics.DrawRectangle(dash, b); }
                    if (s.Type == "Text") { RectangleF h = s.GetResizeHandle(); e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(150, Color.Cyan)), h); e.Graphics.DrawRectangle(Pens.DarkBlue, h.X, h.Y, h.Width, h.Height); }
                }
            }
            e.Graphics.ResetTransform();
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            pb.Focus(); CommitTextEdit();
            Rectangle disp = GetDisplayRect();
            float sX = (float)baseCanvasSize / disp.Width, sY = (float)baseCanvasSize / disp.Height;
            Point imgPt = new Point((int)((e.X - disp.X) * sX), (int)((e.Y - disp.Y) * sY));
            float nx = (float)(e.X - disp.X) / disp.Width, ny = (float)(e.Y - disp.Y) / disp.Height;

            activeDivider = GetDividerAt(nx, ny); if (activeDivider > 0) return;
            if (selectedShape != null && selectedShape.Type == "Text" && selectedShape.GetResizeHandle().Contains(imgPt)) { isResizingText = true; lastMousePos = imgPt; return; }

            selectedShape = null;
            for (int i = shapes.Count - 1; i >= 0; i--) { if (shapes[i].GetBounds().Contains(imgPt)) { selectedShape = shapes[i]; break; } }
            shapes.ForEach(s => s.IsSelected = (s == selectedShape));

            if (selectedShape != null) {
                isDraggingShape = true; lastMousePos = imgPt;
                if (selectedShape.Type == "Text") { textFont = selectedShape.Font; textColor = selectedShape.Color; textBgColor = selectedShape.BgColor; textOpacity = selectedShape.Opacity; cbAlign.SelectedItem = selectedShape.TextAlign; }
                pb.Invalidate(); return;
            }

            if (isTextModeActive) {
                var s = new App_Drawing.DrawShape { Type = "Text", Text = textContent, Font = textFont, Color = textColor, BgColor = textBgColor, BorderColor = textBorderColor, Opacity = textOpacity, TextAlign = cbAlign.SelectedItem.ToString(), TextRect = new RectangleF(imgPt.X, imgPt.Y, 300, 100), IsSelected = true };
                shapes.ForEach(x => x.IsSelected = false); shapes.Add(s); selectedShape = s; isTextModeActive = false; pb.Invalidate(); return;
            }
            if (drawMode != "Select") { drawingShape = new App_Drawing.DrawShape { Type = drawMode, Color = penColor, PenWidth = (int)numPenSize.Value, Start = imgPt, End = imgPt }; shapes.Add(drawingShape); return; }

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
            else pb.Cursor = drawMode != "Select" ? Cursors.Cross : Cursors.Default;

            if (activeDivider > 0) {
                if (activeDivider == 1) div1 = Math.Max(0.1f, Math.Min(0.9f, (cbLayout.SelectedIndex == 1 || cbLayout.SelectedIndex > 2) ? nx : ny));
                else div2 = Math.Max(0.1f, Math.Min(0.9f, (cbLayout.SelectedIndex == 2) ? nx : ny));
                UpdateFramesLayout(); pb.Invalidate();
            } else if (isResizingText && selectedShape != null) {
                selectedShape.TextRect.Width = Math.Max(100, imgPt.X - selectedShape.TextRect.X); selectedShape.TextRect.Height = Math.Max(50, imgPt.Y - selectedShape.TextRect.Y); pb.Invalidate();
            } else if (isDraggingShape && selectedShape != null) {
                selectedShape.Move(imgPt.X - lastMousePos.X, imgPt.Y - lastMousePos.Y); lastMousePos = imgPt; pb.Invalidate();
            } else if (drawingShape != null) { drawingShape.End = imgPt; pb.Invalidate(); }
            else if (isPanningImage && activeFrameIndex >= 0) { frames[activeFrameIndex].OffsetX += (e.X - lastMousePos.X) * sX; frames[activeFrameIndex].OffsetY += (e.Y - lastMousePos.Y) * sY; lastMousePos = e.Location; pb.Invalidate(); }
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) { activeDivider = 0; isDraggingShape = false; isResizingText = false; isPanningImage = false; drawingShape = null; }

        private void Pb_MouseDoubleClick(object sender, MouseEventArgs e) { if (selectedShape != null && selectedShape.Type == "Text") ShowEditBox(selectedShape); }

        private void ShowEditBox(App_Drawing.DrawShape s) {
            editingShape = s; editBox = new TextBox { Multiline = true, Text = s.Text, Font = s.Font };
            if (s.TextAlign == "置中") editBox.TextAlign = HorizontalAlignment.Center; else if (s.TextAlign == "靠右") editBox.TextAlign = HorizontalAlignment.Right; else editBox.TextAlign = HorizontalAlignment.Left;
            Rectangle disp = GetDisplayRect(); float sX = (float)disp.Width / baseCanvasSize, sY = (float)disp.Height / baseCanvasSize;
            RectangleF screenRect = new RectangleF(disp.X + (s.TextRect.X + 10) * sX, disp.Y + (s.TextRect.Y + 10) * sY, (s.TextRect.Width - 20) * sX, (s.TextRect.Height - 20) * sY);
            editBox.Location = Point.Round(screenRect.Location); editBox.Size = Size.Round(screenRect.Size);
            editBox.LostFocus += (sender, ev) => CommitTextEdit();
            pb.Controls.Add(editBox); editBox.BringToFront(); editBox.Focus(); pb.Invalidate();
        }

        private void CommitTextEdit() { if (editBox != null && editingShape != null) { editingShape.Text = editBox.Text; pb.Controls.Remove(editBox); editBox.Dispose(); editBox = null; editingShape = null; pb.Invalidate(); } }

        private void Pb_DragDrop(object sender, DragEventArgs e) {
            Point pt = pb.PointToClient(new Point(e.X, e.Y)); Rectangle disp = GetDisplayRect();
            Point imgPt = new Point((int)((pt.X - disp.X) * (baseCanvasSize / disp.Width)), (int)((pt.Y - disp.Y) * (baseCanvasSize / disp.Height)));
            int dropIndex = -1;
            for (int i = 0; i < frames.Count; i++) { if (GetActualFrameRect(frames[i], baseCanvasSize, baseCanvasSize).Contains(imgPt)) { dropIndex = i; break; } }
            if (dropIndex >= 0) {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && (files[0].ToLower().EndsWith(".jpg") || files[0].ToLower().EndsWith(".png"))) {
                    if (frames[dropIndex].Img != null) frames[dropIndex].Img.Dispose();
                    frames[dropIndex].Img = Image.FromFile(files[0]); frames[dropIndex].Scale = Math.Max(GetActualFrameRect(frames[dropIndex], baseCanvasSize, baseCanvasSize).Width / frames[dropIndex].Img.Width, GetActualFrameRect(frames[dropIndex], baseCanvasSize, baseCanvasSize).Height / frames[dropIndex].Img.Height);
                    activeFrameIndex = dropIndex; UpdateActiveFrameUI(); pb.Invalidate();
                }
            }
        }

        private void SaveImage() {
            CommitTextEdit(); shapes.ForEach(s => s.IsSelected = false); pb.Invalidate();
            Bitmap finalImg = new Bitmap(baseCanvasSize, baseCanvasSize);
            using (Graphics g = Graphics.FromImage(finalImg)) {
                g.FillRectangle(Brushes.White, 0, 0, baseCanvasSize, baseCanvasSize); g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                foreach (var frame in frames) {
                    RectangleF hRect = GetActualFrameRect(frame, baseCanvasSize, baseCanvasSize);
                    g.SetClip(hRect);
                    if (frame.Img != null) { Matrix m = new Matrix(); PointF c = new PointF(hRect.X + hRect.Width / 2 + frame.OffsetX, hRect.Y + hRect.Height / 2 + frame.OffsetY); m.Translate(c.X, c.Y); m.Rotate(frame.Angle); m.Scale(frame.Scale, frame.Scale); m.Translate(-frame.Img.Width / 2f, -frame.Img.Height / 2f); g.Transform = m; g.DrawImage(frame.Img, Point.Empty); g.ResetTransform(); }
                    g.ResetClip();
                }
                g.SmoothingMode = SmoothingMode.AntiAlias;
                foreach (var s in shapes) {
                    if (s.Type == "Text") {
                        using (SolidBrush bg = new SolidBrush(Color.FromArgb(s.Opacity, s.BgColor))) g.FillRectangle(bg, s.TextRect);
                        using (Pen border = new Pen(s.BorderColor, 3)) g.DrawRectangle(border, s.TextRect.X, s.TextRect.Y, s.TextRect.Width, s.TextRect.Height);
                        StringFormat sf = new StringFormat { LineAlignment = StringAlignment.Near };
                        if (s.TextAlign == "置中") sf.Alignment = StringAlignment.Center; else if (s.TextAlign == "靠右") sf.Alignment = StringAlignment.Far; else sf.Alignment = StringAlignment.Near;
                        using (SolidBrush tb = new SolidBrush(s.Color)) g.DrawString(s.Text, s.Font, tb, new RectangleF(s.TextRect.X + 10, s.TextRect.Y + 10, s.TextRect.Width - 20, s.TextRect.Height - 20), sf);
                    } else {
                        using (Pen p = new Pen(s.Color, s.PenWidth)) { int x = Math.Min(s.Start.X, s.End.X), y = Math.Min(s.Start.Y, s.End.Y), w = Math.Abs(s.Start.X - s.End.X), h = Math.Abs(s.Start.Y - s.End.Y); if (s.Type == "Line") g.DrawLine(p, s.Start, s.End); else if (s.Type == "Frame") g.DrawRectangle(p, x, y, w, h); else if (s.Type == "Circle") g.DrawEllipse(p, x, y, w, h); }
                    }
                }
            }
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "Collage_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    EncoderParameters ep = new EncoderParameters(1); ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L); 
                    ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg"); finalImg.Save(sfd.FileName, codec, ep); MessageBox.Show("高解析度拼貼圖儲存成功！");
                }
            }
        }
    }
}
