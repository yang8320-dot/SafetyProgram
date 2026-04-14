/* * 功能：螢幕截圖
 * 對應選單名稱：截圖
 * 對應資料庫名稱：HistoryDB
 * 對應資料表名稱：App_Screenshot
 */
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Threading;

namespace MiniImageStudio {
    public class App_Screenshot : UserControl {
        private MainForm parentForm;
        private PictureBox previewBox;
        private Label statusLabel;
        private static Color AppleBlue = Color.FromArgb(0, 122, 255);
        private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);

        public App_Screenshot(MainForm mainForm) {
            this.parentForm = mainForm;
            this.BackColor = Color.FromArgb(245, 245, 247);
            this.Padding = new Padding(10); [span_7](start_span)//[span_7](end_span)

            // 頂部控制列
            Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 75 };
            
            [span_8](start_span)// 截圖 (藍)[span_8](end_span)
            Button btnNew = new Button() { Text = "截圖", Left = 0, Top = 5, Width = 80, Height = 32, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand };
            btnNew.Click += BtnNew_Click; [span_9](start_span)//[span_9](end_span)

            [span_10](start_span)// 儲存 (灰)[span_10](end_span)
            Button btnSave = new Button() { Text = "儲存", Left = 85, Top = 5, Width = 80, Height = 32, BackColor = Color.Gainsboro, FlatStyle = FlatStyle.Flat, Font = MainFont, Cursor = Cursors.Hand };
            btnSave.Click += BtnSave_Click; [span_11](start_span)//[span_11](end_span)

            statusLabel = new Label() { Text = "點擊「截圖」開始...", Left = 2, Top = 45, AutoSize = true, Font = MainFont, ForeColor = Color.DimGray }; [span_12](start_span)//[span_12](end_span)
            topPanel.Controls.AddRange(new Control[] { btnNew, btnSave, statusLabel });
            
            [span_13](start_span)// 預覽區[span_13](end_span)
            Panel previewContainer = new Panel() { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(5) };
            previewBox = new PictureBox() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(240, 240, 240) }; [span_14](start_span)//[span_14](end_span)
            
            previewContainer.Controls.Add(previewBox);
            this.Controls.Add(previewContainer);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 }); // 間隔
            this.Controls.Add(topPanel);
        }

        private void BtnNew_Click(object sender, EventArgs e) {
            parentForm.Hide(); [span_15](start_span)//[span_15](end_span)
            Application.DoEvents();
            Thread.Sleep(250); [span_16](start_span)//[span_16](end_span)

            using (SnippingOverlayForm snipForm = new SnippingOverlayForm()) {
                if (snipForm.ShowDialog() == DialogResult.OK && snipForm.ResultImage != null) {
                    if (previewBox.Image != null) previewBox.Image.Dispose();
                    previewBox.Image = (Image)snipForm.ResultImage.Clone(); [span_17](start_span)//[span_17](end_span)
                    Clipboard.SetImage(previewBox.Image);
                    statusLabel.Text = "截圖成功！已複製。";
                    statusLabel.ForeColor = Color.FromArgb(0, 153, 76);
                    App_History.WriteLog("Screenshot|Capture Success");
                } else {
                    statusLabel.Text = "已取消截圖。"; [span_18](start_span)//[span_18](end_span)
                }
            }
            parentForm.Show();
        }

        private void BtnSave_Click(object sender, EventArgs e) {
            if (previewBox.Image == null) return; [span_19](start_span)//[span_19](end_span)
            [span_20](start_span)using (SaveFileDialog sfd = new SaveFileDialog()) { //[span_20](end_span)
                sfd.Filter = "PNG 圖片|*.png|JPEG 圖片|*.jpg";
                sfd.FileName = "截圖_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"); [span_21](start_span)//[span_21](end_span)
                if (sfd.ShowDialog() == DialogResult.OK) {
                    ImageFormat format = sfd.FileName.EndsWith(".jpg") ? ImageFormat.Jpeg : ImageFormat.Png; [span_22](start_span)//[span_22](end_span)
                    previewBox.Image.Save(sfd.FileName, format);
                    App_History.WriteLog($"Screenshot|Saved {sfd.FileName}");
                }
            }
        }
    }

    // 核心：全螢幕截圖覆蓋視窗 (Overlay)
    public class SnippingOverlayForm : Form {
        private Bitmap screenBmp;
        private Point startPt; [span_23](start_span)//[span_23](end_span)
        private Rectangle selectRect;
        private bool isDragging = false;
        public Image ResultImage { get; private set; [span_24](start_span)} //[span_24](end_span)

        public SnippingOverlayForm() {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true; this.ShowInTaskbar = false; [span_25](start_span)//[span_25](end_span)
            this.Cursor = Cursors.Cross; this.DoubleBuffered = true; 
            this.Location = SystemInformation.VirtualScreen.Location;
            this.Size = SystemInformation.VirtualScreen.Size;
            screenBmp = new Bitmap(this.Width, this.Height); [span_26](start_span)//[span_26](end_span)
            using (Graphics g = Graphics.FromImage(screenBmp)) {
                g.CopyFromScreen(this.Location, Point.Empty, this.Size); [span_27](start_span)//[span_27](end_span)
            }
        }

        protected override void OnPaint(PaintEventArgs e) {
            e.Graphics.DrawImageUnscaled(screenBmp, 0, 0); [span_28](start_span)//[span_28](end_span)
            using (SolidBrush dimBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0))) {
                e.Graphics.FillRectangle(dimBrush, this.ClientRectangle); [span_29](start_span)//[span_29](end_span)
            }
            if (selectRect.Width > 0 && selectRect.Height > 0) {
                e.Graphics.DrawImage(screenBmp, selectRect, selectRect, GraphicsUnit.Pixel); [span_30](start_span)//[span_30](end_span)
                using (Pen borderPen = new Pen(Color.Red, 2)) {
                    e.Graphics.DrawRectangle(borderPen, selectRect); [span_31](start_span)//[span_31](end_span)
                }
            }
        }

        protected override void OnMouseDown(MouseEventArgs e) {
            if (e.Button == MouseButtons.Left) { isDragging = true; startPt = e.Location; selectRect = new Rectangle(e.X, e.Y, 0, 0); [span_32](start_span)} //[span_32](end_span)
            else if (e.Button == MouseButtons.Right) { this.DialogResult = DialogResult.Cancel; this.Close(); [span_33](start_span)} //[span_33](end_span)
        }

        protected override void OnMouseMove(MouseEventArgs e) {
            if (isDragging) {
                int x = Math.Min(startPt.X, e.X); int y = Math.Min(startPt.Y, e.Y); [span_34](start_span)//[span_34](end_span)
                selectRect = new Rectangle(x, y, Math.Abs(startPt.X - e.X), Math.Abs(startPt.Y - e.Y)); [span_35](start_span)//[span_35](end_span)
                this.Invalidate(); 
            }
        }

        protected override void OnMouseUp(MouseEventArgs e) {
            if (isDragging) {
                isDragging = false;
                [span_36](start_span)if (selectRect.Width > 5 && selectRect.Height > 5) { //[span_36](end_span)
                    ResultImage = new Bitmap(selectRect.Width, selectRect.Height);
                    using (Graphics g = Graphics.FromImage(ResultImage)) {
                        g.DrawImage(screenBmp, new Rectangle(0, 0, selectRect.Width, selectRect.Height), selectRect, GraphicsUnit.Pixel); [span_37](start_span)//[span_37](end_span)
                    }
                    this.DialogResult = DialogResult.OK;
                } else { this.DialogResult = DialogResult.Cancel; [span_38](start_span)} //[span_38](end_span)
                this.Close(); [span_39](start_span)//[span_39](end_span)
            }
        }
    }
}
