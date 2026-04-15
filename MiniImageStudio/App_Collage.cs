/* * 功能：圖片拼貼與文字註解 (包含底框顏色)
 * 對應選單名稱：拼貼
 * 對應資料庫名稱：HistoryDB
 * 對應資料表名稱：App_Collage
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class App_Collage : UserControl {
        private PictureBox pb;
        private Bitmap img1, img2;
        private TextBox txtAnnotation;

        public App_Collage() {
            this.Padding = new Padding(10);
            
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            
            Button btnLoad1 = new Button { Text = "載入左圖", Width = 80, Left = 0, FlatStyle = FlatStyle.Flat };
            Button btnLoad2 = new Button { Text = "載入右圖", Width = 80, Left = 85, FlatStyle = FlatStyle.Flat };
            Button btnProcess = new Button { Text = "產生拼貼", Width = 80, Left = 170, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            
            Label lblText = new Label { Text = "註解:", Left = 260, Top = 8, AutoSize = true };
            txtAnnotation = new TextBox { Left = 300, Top = 5, Width = 120, Text = "自訂文字" };
            
            Button btnSave = new Button { Text = "儲存", Width = 60, Left = 430, FlatStyle = FlatStyle.Flat };

            btnLoad1.Click += (s, e) => img1 = LoadImage("左圖");
            btnLoad2.Click += (s, e) => img2 = LoadImage("右圖");
            btnProcess.Click += (s, e) => GenerateCollage();
            btnSave.Click += (s, e) => SaveImage();

            topPanel.Controls.AddRange(new Control[] { btnLoad1, btnLoad2, btnProcess, lblText, txtAnnotation, btnSave });

            pb = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(240, 240, 240) };
            
            this.Controls.Add(pb);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(topPanel);
        }

        private Bitmap LoadImage(string side) {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Images|*.jpg;*.png", Title = $"請選擇{side}" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    return new Bitmap(ofd.FileName);
                }
            }
            return null;
        }

        private void GenerateCollage() {
            if (img1 == null || img2 == null) {
                MessageBox.Show("請先載入左圖與右圖！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 計算拼貼後的畫布大小 (左右並排，高度取最高者)
            int totalWidth = img1.Width + img2.Width;
            int maxHeight = Math.Max(img1.Height, img2.Height);
            
            Bitmap canvas = new Bitmap(totalWidth, maxHeight);
            
            using (Graphics g = Graphics.FromImage(canvas)) {
                g.Clear(Color.White);
                
                // 畫上左右兩張圖
                g.DrawImage(img1, 0, 0, img1.Width, img1.Height);
                g.DrawImage(img2, img1.Width, 0, img2.Width, img2.Height);

                // 畫上註解文字 (包含文字底框底色)
                string text = txtAnnotation.Text;
                if (!string.IsNullOrEmpty(text)) {
                    Font font = new Font("Microsoft JhengHei UI", 36, FontStyle.Bold);
                    SizeF textSize = g.MeasureString(text, font);
                    
                    // 設定底框位置 (左上角留白 20px)
                    RectangleF bgRect = new RectangleF(20, 20, textSize.Width + 20, textSize.Height + 20);
                    
                    // 畫半透明黑色底框
                    using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(180, 0, 0, 0))) {
                        g.FillRectangle(bgBrush, bgRect);
                    }
                    
                    // 畫白色文字
                    g.DrawString(text, font, Brushes.White, 30, 30);
                }
            }

            if (pb.Image != null) pb.Image.Dispose();
            pb.Image = canvas;
            App_History.WriteLog("Collage|Generated Successfully");
        }

        private void SaveImage() {
            if (pb.Image == null) return;
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "JPEG|*.jpg", FileName = "拼貼_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    pb.Image.Save(sfd.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                    App_History.WriteLog("Collage|Saved");
                }
            }
        }
    }
}
