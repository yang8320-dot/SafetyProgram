/* * 功能：批次壓縮與縮圖 (支援單/多筆拖曳上傳，自動切換至下載資料夾)
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
        private string customOutputDir = "";

        public App_Compress() {
            this.Font = MainForm.UI_Font;
            this.Padding = new Padding(10);
            InitializeUI();
        }

        private void InitializeUI() {
            FlowLayoutPanel topPanel = new FlowLayoutPanel { 
                Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, 
                Padding = new Padding(5), BackColor = SystemColors.Control 
            };
            
            Button btnAdd = new Button { Text = "加入圖片", Width = 110, Height = 40, Margin = new Padding(5) };
            Label lblSize = new Label { Text = "最長邊像素:", AutoSize = true, Margin = new Padding(5, 15, 0, 5) };
            txtMaxSize = new TextBox { Text = "1024", Width = 80, Margin = new Padding(5, 12, 15, 5) };
            Button btnSetFolder = new Button { Text = "設定輸出資料夾", Width = 150, Height = 40, Margin = new Padding(5) };
            Button btnProcess = new Button { Text = "開始批次處理", Width = 140, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Margin = new Padding(5) };
            
            lblStatus = new Label { Text = "等待中...", AutoSize = true, Margin = new Padding(15, 15, 5, 5) };
            
            FlowLayoutPanel infoPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, BackColor = SystemColors.Control };
            lblOutPath = new Label { Text = "輸出至: (預設為原資料夾內的 Compressed 子目錄)", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(10, 0, 10, 10) };
            infoPanel.Controls.Add(lblOutPath);

            btnAdd.Click += BtnAdd_Click;
            btnSetFolder.Click += BtnSetFolder_Click;
            btnProcess.Click += async (s, e) => await ProcessImagesAsync();

            topPanel.Controls.AddRange(new Control[] { btnAdd, lblSize, txtMaxSize, btnSetFolder, btnProcess, lblStatus });
            
            listFiles = new ListBox { Dock = DockStyle.Fill, ItemHeight = 20, AllowDrop = true }; // 啟用拖曳
            listFiles.DragEnter += ListFiles_DragEnter;
            listFiles.DragDrop += ListFiles_DragDrop;
            
            this.Controls.Add(listFiles);
            this.Controls.Add(infoPanel);
            this.Controls.Add(topPanel);
        }

        private void BtnAdd_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "Images|*.jpg;*.png;*.jpeg;*.bmp" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    foreach (var file in ofd.FileNames) {
                        if (!listFiles.Items.Contains(file)) listFiles.Items.Add(file);
                    }
                    lblStatus.Text = $"已加入 {listFiles.Items.Count} 張圖片";
                }
            }
        }

        // --- 處理拖曳事件 ---
        private void ListFiles_DragEnter(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void ListFiles_DragDrop(object sender, DragEventArgs e) {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            bool hasNewFile = false;

            foreach (string file in files) {
                string ext = Path.GetExtension(file).ToLower();
                if (ext == ".jpg" || ext == ".png" || ext == ".jpeg" || ext == ".bmp") {
                    if (!listFiles.Items.Contains(file)) {
                        listFiles.Items.Add(file);
                        hasNewFile = true;
                    }
                }
            }

            if (hasNewFile) {
                lblStatus.Text = $"已加入 {listFiles.Items.Count} 張圖片";
                
                // 拖曳上傳時，自動將預設存檔路徑設為「下載 (Downloads) / Compressed」
                string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Compressed");
                customOutputDir = downloadsPath;
                lblOutPath.Text = $"輸出至: {customOutputDir}";
                lblOutPath.ForeColor = Color.Blue;
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
            lblStatus.Text = "處理中，請稍候..."; lblStatus.ForeColor = Color.OrangeRed;
            var files = listFiles.Items.Cast<string>().ToList();

            await Task.Run(() => {
                foreach (var file in files) {
                    try {
                        using (Image img = Image.FromFile(file)) {
                            int newW = img.Width, newH = img.Height;
                            if (Math.Max(img.Width, img.Height) > maxSize) {
                                float ratio = (float)maxSize / Math.Max(img.Width, img.Height);
                                newW = (int)(img.Width * ratio); newH = (int)(img.Height * ratio);
                            }
                            string targetDir = customOutputDir;
                            if (string.IsNullOrEmpty(targetDir)) targetDir = Path.Combine(Path.GetDirectoryName(file), "Compressed");
                            if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);
                            string outPath = Path.Combine(targetDir, Path.GetFileNameWithoutExtension(file) + "_comp.jpg");

                            using (Bitmap newImg = new Bitmap(newW, newH)) {
                                using (Graphics g = Graphics.FromImage(newImg)) {
                                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    g.DrawImage(img, 0, 0, newW, newH);
                                }
                                EncoderParameters ep = new EncoderParameters(1);
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 85L);
                                ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().First(c => c.MimeType == "image/jpeg");
                                newImg.Save(outPath, codec, ep);
                            }
                        }
                    } catch { }
                }
            });
            lblStatus.Text = "全部處理完成！"; lblStatus.ForeColor = Color.SeaGreen;
            listFiles.Items.Clear();
        }
    }
}
