/* * 功能：圖片繪製與註解
 * 對應選單名稱：繪製
 * 對應資料庫名稱：HistoryDB
 * 對應資料表名稱：App_Drawing
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class App_Drawing : UserControl {
        private PictureBox pb;
        private Bitmap canvas;
        private Color currentColor = Color.Red;
        private string currentMode = "Line"; // 模式: Line, Frame, Text
        private Point startPoint;
        private bool isDrawing = false;
        private TextBox txtAnnotation;

        public App_Drawing() {
            this.Padding = new Padding(10);
            
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            Button btnLoad = new Button { Text = "載入", Width = 60, Left = 0, FlatStyle = FlatStyle.Flat };
            Button btnColor = new Button { Text = "顏色", Width = 60, Left = 65, FlatStyle = FlatStyle.Flat, BackColor = currentColor };
            Button btnLine = new Button { Text = "畫線", Width = 60, Left = 130, FlatStyle = FlatStyle.Flat };
            Button btnFrame = new Button { Text = "畫框", Width = 60, Left = 195, FlatStyle = FlatStyle.Flat };
            Button btnText = new Button { Text = "註解", Width = 60, Left = 260, FlatStyle = FlatStyle.Flat };
            Button btnSave = new Button { Text = "儲存", Width = 60, Left = 325, FlatStyle = FlatStyle.Flat };
            
            txtAnnotation = new TextBox { Left = 390, Top = 5, Width = 150, Text = "輸入註解文字" };

            btnLoad.Click += (s, e) => LoadImage();
            btnColor.Click += (s, e) => {
                using (ColorDialog cd = new ColorDialog()) {
                    if (cd.ShowDialog() == DialogResult.OK) { currentColor = cd.Color; btnColor.BackColor = currentColor; }
                }
            };
            btnLine.Click += (s, e) => currentMode = "Line";
            btnFrame.Click += (s, e) => currentMode = "Frame";
            btnText.Click += (s, e) => currentMode = "Text";
            btnSave.Click += (s, e) => SaveImage();

            topPanel.Controls.AddRange(new Control[] { btnLoad, btnColor, btnLine, btnFrame, btnText, btnSave, txtAnnotation });

            pb = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White, Cursor = Cursors.Cross };
            pb.MouseDown += Pb_MouseDown;
            pb.MouseUp += Pb_MouseUp;

            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(topPanel);
        }

        private void LoadImage() {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Images|*.png;*.jpg;*.jpeg" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    canvas = new Bitmap(ofd.FileName);
                    pb.Image = canvas;
                }
            }
        }

        private void Pb_MouseDown(object sender, MouseEventArgs e) {
            if (canvas == null) return;
            isDrawing = true;
            startPoint = GetImagePoint(e.Location);
        }

        private void Pb_MouseUp(object sender, MouseEventArgs e) {
            if (!isDrawing || canvas == null) return;
            isDrawing = false;
            Point endPoint = GetImagePoint(e.Location);
            
            using (Graphics g = Graphics.FromImage(canvas)) {
                using (Pen pen = new Pen(currentColor, 3)) {
                    if (currentMode == "Line") {
                        g.DrawLine(pen, startPoint, endPoint);
                    } else if (currentMode == "Frame") {
                        int x = Math.Min(startPoint.X, endPoint.X);
                        int y = Math.Min(startPoint.Y, endPoint.Y);
                        g.DrawRectangle(pen, x, y, Math.Abs(startPoint.X - endPoint.X), Math.Abs(startPoint.Y - endPoint.Y));
                    } else if (currentMode == "Text") {
                        Font f = new Font("Microsoft JhengHei UI", 24, FontStyle.Bold);
                        g.DrawString(txtAnnotation.Text, f, new SolidBrush(currentColor), endPoint);
                    }
                }
            }
            pb.Invalidate();
            App_History.WriteLog($"Drawing|Mode:{currentMode}");
        }

        // 將 PictureBox 座標轉換為圖片實際座標
        private Point GetImagePoint(Point pbPoint) {
            if (pb.Image == null) return Point.Empty;
            float ratioX = (float)pb.Image.Width / pb.Width;
            float ratioY = (float)pb.Image.Height / pb.Height;
            float ratio = Math.Max(ratioX, ratioY);
            int imgW = (int)(pb.Width * ratio);
            int imgH = (int)(pb.Height * ratio);
            int offX = (pb.Width - imgW) / 2;
            int offY = (pb.Height - imgH) / 2;
            int x = (int)((pbPoint.X - offX) * ratioX);
            int y = (int)((pbPoint.Y - offY) * ratioY);
            return new Point(x, y);
        }

        private void SaveImage() {
            if (canvas == null) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg" }) {
                if (sfd.ShowDialog() == DialogResult.OK) canvas.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }
    }
}
