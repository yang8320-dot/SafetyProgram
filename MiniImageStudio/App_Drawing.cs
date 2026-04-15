/* * 功能：繪製模組 (修復座標偏移，加入樣板功能)
 */
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class App_Drawing : UserControl {
        private PictureBox pb;
        private Bitmap canvas;
        private Color currentColor = Color.Red;
        private string currentMode = "Line";
        private int penWidth = 3;
        private Point startPoint;
        private bool isDrawing = false;
        private TextBox txtInput;

        public App_Drawing() {
            this.Font = MainForm.UI_Font;
            InitializeUI();
        }

        private void InitializeUI() {
            Panel ctrlPanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = SystemColors.Control };
            
            Button btnLoad = new Button { Text = "載入圖片", Left = 10, Top = 10, Width = 90 };
            Button btnColor = new Button { Text = "顏色選擇", Left = 105, Top = 10, Width = 90, BackColor = currentColor };
            
            // 模式選擇樣板
            ComboBox cbMode = new ComboBox { Left = 200, Top = 15, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cbMode.Items.AddRange(new string[] { "自由畫線", "標準矩形", "註解文字" });
            cbMode.SelectedIndex = 0;

            ComboBox cbSize = new ComboBox { Left = 310, Top = 15, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbSize.Items.AddRange(new string[] { "細(2pt)", "中(5pt)", "粗(10pt)" });
            cbSize.SelectedIndex = 0;

            txtInput = new TextBox { Left = 400, Top = 15, Width = 150, Text = "在此輸入註解內容" };
            Button btnSave = new Button { Text = "儲存圖片", Left = 560, Top = 10, Width = 90 };

            btnLoad.Click += (s, e) => LoadImage();
            btnColor.Click += (s, e) => {
                using (ColorDialog cd = new ColorDialog()) {
                    if (cd.ShowDialog() == DialogResult.OK) {
                        currentColor = cd.Color;
                        btnColor.BackColor = currentColor;
                    }
                }
            };
            cbMode.SelectedIndexChanged += (s, e) => {
                if (cbMode.SelectedIndex == 0) currentMode = "Line";
                else if (cbMode.SelectedIndex == 1) currentMode = "Frame";
                else currentMode = "Text";
            };
            cbSize.SelectedIndexChanged += (s, e) => {
                penWidth = cbSize.SelectedIndex == 0 ? 2 : (cbSize.SelectedIndex == 1 ? 5 : 10);
            };
            btnSave.Click += (s, e) => SaveImage();

            ctrlPanel.Controls.AddRange(new Control[] { btnLoad, btnColor, cbMode, cbSize, txtInput, btnSave });

            pb = new PictureBox { 
                Dock = DockStyle.Fill, 
                SizeMode = PictureBoxSizeMode.Zoom, 
                BackColor = Color.DarkGray,
                BorderStyle = BorderStyle.Fixed3D 
            };
            
            pb.MouseDown += Pb_MouseDown;
            pb.MouseMove += Pb_MouseMove;
            pb.MouseUp += Pb_MouseUp;

            this.Controls.Add(pb);
            this.Controls.Add(ctrlPanel);
        }

        private void LoadImage() {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Image Files|*.jpg;*.png;*.bmp" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    canvas = new Bitmap(ofd.FileName);
                    pb.Image = canvas;
                }
            }
        }

        // 核心修正：精確轉換 PictureBox 座標到 Bitmap 座標
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

            return new Point(
                (int)((p.X - offsetX) / scale),
                (int)((p.Y - offsetY) / scale)
            );
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            if (canvas == null) return;
            isDrawing = true;
            startPoint = TranslatePoint(e.Location);
        }

        private void Pb_MouseMove(object sender, MouseEventArgs e) {
            if (isDrawing) pb.Invalidate(); // 觸發重繪預覽
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            if (!isDrawing || canvas == null) return;
            isDrawing = false;
            Point endPoint = TranslatePoint(e.Location);

            using (Graphics g = Graphics.FromImage(canvas)) {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (Pen p = new Pen(currentColor, penWidth)) {
                    if (currentMode == "Line") g.DrawLine(p, startPoint, endPoint);
                    else if (currentMode == "Frame") {
                        int x = Math.Min(startPoint.X, endPoint.X);
                        int y = Math.Min(startPoint.Y, endPoint.Y);
                        g.DrawRectangle(p, x, y, Math.Abs(startPoint.X - endPoint.X), Math.Abs(startPoint.Y - endPoint.Y));
                    } else if (currentMode == "Text") {
                        g.DrawString(txtInput.Text, new Font(MainForm.UI_Font.FontFamily, penWidth * 5), new SolidBrush(currentColor), endPoint);
                    }
                }
            }
            pb.Refresh();
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
