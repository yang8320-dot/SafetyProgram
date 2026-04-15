/* * 功能：批次壓縮與縮圖 (含子資料夾建立與自訂路徑)
 * 對應選單名稱：壓縮
 * 對應資料庫名稱：HistoryDB
 * 對應資料表名稱：App_Compress
 */
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniImageStudio {
    public class App_Compress : UserControl {
        private TextBox txtMaxSize;
        private ListBox listFiles;
        private Label lblStatus, lblOutPath;
        private string customOutputDir = ""; // 空白代表使用預設原資料夾內的 Compressed

        public App_Compress() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
        }

        private void InitializeUI() {
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 90, BackColor = SystemColors.Control };
            
            Button btnAdd = new Button { Text = "加入圖片", Width = 90, Left = 10, Top = 10 };
            Label lblSize = new Label { Text = "最長邊像素:", Left = 110, Top = 15, AutoSize = true };
            txtMaxSize = new TextBox { Text = "1024", Left = 195, Top = 12, Width = 60 };
            
            Button btnSetFolder = new Button { Text = "設定輸出資料夾", Width = 130, Left = 270, Top = 10 };
            lblOutPath = new Label { Text = "輸出至: (預設為原資料夾內的 Compressed 子目錄)", Left = 10, Top = 55, AutoSize = true, ForeColor = Color.DimGray };

            Button btnProcess = new Button { Text = "開始批次處理", Width = 120, Left = 410, Top = 10, BackColor = Color.SteelBlue, ForeColor = Color.White };
            lblStatus = new Label { Text = "等待中...", Left = 540, Top = 15, AutoSize = true };

            btnAdd.Click += BtnAdd_Click;
            btnSetFolder.Click += BtnSetFolder_Click;
            btnProcess.Click += async (s, e) => await ProcessImagesAsync();

            topPanel.Controls.AddRange(new Control[] { btnAdd, lblSize, txtMaxSize, btnSetFolder, lblOutPath, btnProcess, lblStatus });

            listFiles = new ListBox { Dock = DockStyle.Fill, ItemHeight = 20 };
            
            this.Controls.Add(listFiles);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 });
            this.Controls.Add(topPanel);
        }

        private void BtnAdd_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "Images|*.jpg;*.png;*.jpeg" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    foreach (var file in ofd.FileNames) {
                        if (!listFiles.Items.Contains(file)) listFiles.Items.Add(file);
                    }
                    lblStatus.Text = $"已加入 {listFiles.Items.Count} 張圖片";
                }
            }
        }

        private void BtnSetFolder_Click(object sender, EventArgs e) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "請選擇壓縮後的圖片存放資料夾" }) {
                if (fbd.ShowDialog() == DialogResult.OK) {
                    customOutputDir = fbd.SelectedPath;
                    lblOutPath.Text = $"輸出至: {customOutputDir}";
                    lblOutPath.ForeColor = Color.Blue;
                }
            }
        }

        private async Task ProcessImagesAsync() {
            if (listFiles.Items.Count == 0) return;
            if (!int.TryParse(txtMaxSize.Text, out int maxSize)) maxSize = 1024;
            
            lblStatus.Text = "處理中，請稍候...";
            lblStatus.ForeColor = Color.OrangeRed;
            var files = listFiles.Items.Cast<string>().ToList();

            // 背景處理避免畫面凍結 (Thread-Safety)
            await Task.Run(() => {
                foreach (var file in files) {
                    try {
                        using (Image img = Image.FromFile(file)) {
                            // 計算等比例
                            int newW = img.Width, newH = img.Height;
                            if (Math.Max(img.Width, img.Height) > maxSize) {
                                float ratio = (float)maxSize / Math.Max(img.Width, img.Height);
                                newW = (int)(img.Width * ratio);
                                newH = (int)(img.Height * ratio);
                            }

                            // 決定輸出路徑
                            string targetDir = customOutputDir;
                            if (string.IsNullOrEmpty(targetDir)) {
                                targetDir = Path.Combine(Path.GetDirectoryName(file), "Compressed");
                            }
                            if (!Directory.Exists(targetDir)) {
                                Directory.CreateDirectory(targetDir);
                            }

                            string outPath = Path.Combine(targetDir, Path.GetFileNameWithoutExtension(file) + "_comp.jpg");

                            // 繪製新尺寸並壓縮存檔
                            using (Bitmap newImg = new Bitmap(newW, newH)) {
                                using (Graphics g = Graphics.FromImage(newImg)) {
                                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    g.DrawImage(img, 0, 0, newW, newH);
                                }
                                
                                EncoderParameters ep = new EncoderParameters(1);
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 85L); // 85% 品質
                                ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                                
                                newImg.Save(outPath, codec, ep);
                            }
                        }
                    } catch { /* 若某檔案損壞則略過，繼續處理下一張 */ }
                }
            });

            lblStatus.Text = "全部處理完成！";
            lblStatus.ForeColor = Color.SeaGreen;
            listFiles.Items.Clear();
        }
    }
}
