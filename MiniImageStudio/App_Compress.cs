/* * 功能：圖片壓縮與批次等比例縮圖
 * 對應選單名稱：壓縮
 * 對應資料庫名稱：HistoryDB
 * 對應資料表名稱：App_Compress
 */
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Linq;

namespace MiniImageStudio {
    public class App_Compress : UserControl {
        private TextBox txtMaxSize;
        private ListBox listFiles;
        private Label lblStatus;

        public App_Compress() {
            this.Padding = new Padding(10);

            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            Button btnAdd = new Button { Text = "加入圖片", Width = 80, Left = 0, FlatStyle = FlatStyle.Flat };
            Label lblSize = new Label { Text = "最長邊像素:", Left = 90, Top = 8, AutoSize = true };
            txtMaxSize = new TextBox { Text = "1024", Left = 170, Top = 5, Width = 60 };
            Button btnProcess = new Button { Text = "開始批次壓縮", Width = 100, Left = 240, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White };
            lblStatus = new Label { Text = "等待中...", Left = 350, Top = 8, AutoSize = true };

            btnAdd.Click += BtnAdd_Click;
            btnProcess.Click += async (s, e) => await ProcessImagesAsync();

            topPanel.Controls.AddRange(new Control[] { btnAdd, lblSize, txtMaxSize, btnProcess, lblStatus });

            listFiles = new ListBox { Dock = DockStyle.Fill };
            
            this.Controls.Add(listFiles);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(topPanel);
        }

        private void BtnAdd_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "Images|*.jpg;*.png" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    foreach (var file in ofd.FileNames) listFiles.Items.Add(file);
                }
            }
        }

        private async Task ProcessImagesAsync() {
            if (listFiles.Items.Count == 0) return;
            if (!int.TryParse(txtMaxSize.Text, out int maxSize)) maxSize = 1024;
            
            lblStatus.Text = "處理中...";
            var files = listFiles.Items.Cast<string>().ToList();

            // 實作背景批次處理
            await Task.Run(() => {
                foreach (var file in files) {
                    try {
                        using (Image img = Image.FromFile(file)) {
                            // 計算等比例縮放
                            int newW = img.Width;
                            int newH = img.Height;
                            if (Math.Max(img.Width, img.Height) > maxSize) {
                                float ratio = (float)maxSize / Math.Max(img.Width, img.Height);
                                newW = (int)(img.Width * ratio);
                                newH = (int)(img.Height * ratio);
                            }

                            using (Bitmap newImg = new Bitmap(img, new Size(newW, newH))) {
                                string outPath = Path.Combine(Path.GetDirectoryName(file), "Compressed_" + Path.GetFileName(file));
                                
                                // 設定 JPEG 高品質壓縮
                                EncoderParameters ep = new EncoderParameters(1);
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 85L);
                                ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                                
                                newImg.Save(outPath, codec, ep);
                            }
                        }
                    } catch { /* 略過損壞檔案 */ }
                }
            });

            lblStatus.Text = "完成！檔案已儲存於原資料夾";
            App_History.WriteLog($"Compress|Batch Processed {files.Count} files");
            listFiles.Items.Clear();
        }
    }
}
