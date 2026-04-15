/* * 功能：批次壓縮與縮圖 (清單 65%，預覽 35%，新增開啟輸出資料夾功能)
 */
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics; // 用於開啟資料夾

namespace MiniImageStudio {
    public class App_Compress : UserControl {
        private TextBox txtMaxSize;
        private ListBox listFiles;
        private PictureBox pbPreview; 
        private Label lblStatus, lblOutPath;
        private SplitContainer splitContainer; 
        private string customOutputDir = "";
        private bool isSplitterSet = false;

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
            
            Button btnAdd = new Button { Text = "加入圖片", Width = 90, Height = 40, Margin = new Padding(5) };
            Button btnRemove = new Button { Text = "移除選取", Width = 90, Height = 40, Margin = new Padding(5), BackColor = Color.IndianRed, ForeColor = Color.White };
            
            Label lblSize = new Label { Text = "最長邊像素:", AutoSize = true, Margin = new Padding(10, 15, 0, 5) };
            txtMaxSize = new TextBox { Text = "1024", Width = 70, Margin = new Padding(5, 12, 15, 5) };
            
            Button btnSetFolder = new Button { Text = "設定輸出資料夾", Width = 130, Height = 40, Margin = new Padding(5) };
            Button btnOpenFolder = new Button { Text = "開啟輸出資料夾", Width = 130, Height = 40, Margin = new Padding(5), BackColor = Color.DarkOrange, ForeColor = Color.White };
            Button btnProcess = new Button { Text = "開始批次處理", Width = 120, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Margin = new Padding(5) };
            
            lblStatus = new Label { Text = "等待中...", AutoSize = true, Margin = new Padding(15, 15, 5, 5) };
            
            FlowLayoutPanel infoPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, BackColor = SystemColors.Control };
            lblOutPath = new Label { Text = "輸出至: (預設為原資料夾內的 Compressed 子目錄)", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(10, 0, 10, 10) };
            infoPanel.Controls.Add(lblOutPath);

            splitContainer = new SplitContainer { Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
            
            listFiles = new ListBox { Dock = DockStyle.Fill, ItemHeight = 20, AllowDrop = true };
            listFiles.DragEnter += ListFiles_DragEnter;
            listFiles.DragDrop += ListFiles_DragDrop;
            listFiles.SelectedIndexChanged += ListFiles_SelectedIndexChanged;
            listFiles.KeyDown += ListFiles_KeyDown; 

            pbPreview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.WhiteSmoke, BorderStyle = BorderStyle.Fixed3D };

            splitContainer.Panel1.Controls.Add(listFiles);
            splitContainer.Panel2.Controls.Add(pbPreview);

            btnAdd.Click += BtnAdd_Click;
            btnRemove.Click += (s, e) => RemoveSelectedFile();
            btnSetFolder.Click += BtnSetFolder_Click;
            btnOpenFolder.Click += BtnOpenFolder_Click; // 綁定開啟資料夾事件
            btnProcess.Click += async (s, e) => await ProcessImagesAsync();

            topPanel.Controls.AddRange(new Control[] { btnAdd, btnRemove, lblSize, txtMaxSize, btnSetFolder, btnOpenFolder, btnProcess, lblStatus });
            
            this.Controls.Add(splitContainer);
            this.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 10 }); 
            this.Controls.Add(infoPanel);
            this.Controls.Add(topPanel);

            // --- 精準的 65% / 35% 比例計算 ---
            this.Resize += (s, e) => {
                if (!isSplitterSet && this.Width > 200) {
                    splitContainer.SplitterDistance = (int)(this.Width * 0.65); // 設定左邊佔 65%
                    isSplitterSet = true;
                }
            };
        }

        private void BtnAdd_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "Images|*.jpg;*.png;*.jpeg;*.bmp" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    bool hasNewFile = false;
                    foreach (var file in ofd.FileNames) {
                        if (!listFiles.Items.Contains(file)) {
                            listFiles.Items.Add(file);
                            hasNewFile = true;
                        }
                    }
                    if (hasNewFile) UpdateStatusAndPath();
                }
            }
        }

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
                UpdateStatusAndPath();
            }
        }

        private void UpdateStatusAndPath() {
            lblStatus.Text = $"已加入 {listFiles.Items.Count} 張圖片";
            if (string.IsNullOrEmpty(customOutputDir)) {
                string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Compressed");
                customOutputDir = downloadsPath;
                lblOutPath.Text = $"輸出至: {customOutputDir}";
                lblOutPath.ForeColor = Color.Blue;
            }
        }

        private void RemoveSelectedFile() {
            if (listFiles.SelectedIndex >= 0) {
                int selectedIndex = listFiles.SelectedIndex;
                listFiles.Items.RemoveAt(selectedIndex);
                if (listFiles.Items.Count > 0) {
                    listFiles.SelectedIndex = Math.Min(selectedIndex, listFiles.Items.Count - 1);
                } else {
                    ClearPreview();
                }
                lblStatus.Text = $"已加入 {listFiles.Items.Count} 張圖片";
            }
        }

        private void ListFiles_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Delete) {
                RemoveSelectedFile();
            }
        }

        private void ListFiles_SelectedIndexChanged(object sender, EventArgs e) {
            if (listFiles.SelectedItem != null) {
                string filePath = listFiles.SelectedItem.ToString();
                try {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read)) {
                        using (MemoryStream ms = new MemoryStream()) {
                            fs.CopyTo(ms);
                            if (pbPreview.Image != null) pbPreview.Image.Dispose();
                            pbPreview.Image = Image.FromStream(ms);
                        }
                    }
                } catch {
                    ClearPreview();
                }
            }
        }

        private void ClearPreview() {
            if (pbPreview.Image != null) {
                pbPreview.Image.Dispose();
                pbPreview.Image = null;
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

        // --- 開啟輸出資料夾功能 ---
        private void BtnOpenFolder_Click(object sender, EventArgs e) {
            string pathToOpen = customOutputDir;
            
            // 如果尚未設定路徑，則使用預設的下載資料夾路徑
            if (string.IsNullOrEmpty(pathToOpen)) {
                pathToOpen = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Compressed");
            }
            
            // 如果資料夾還不存在，先自動建立，避免開啟失敗報錯
            if (!Directory.Exists(pathToOpen)) {
                try {
                    Directory.CreateDirectory(pathToOpen);
                } catch {
                    MessageBox.Show("無法建立或開啟資料夾路徑！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            
            // 呼叫系統檔案總管開啟該路徑
            Process.Start("explorer.exe", pathToOpen);
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
            ClearPreview();
        }
    }
}
