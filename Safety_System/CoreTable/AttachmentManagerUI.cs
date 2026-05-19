/// FILE: Safety_System/CoreTable/AttachmentManagerUI.cs ///
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class AttachmentManagerUI : Form
    {
        public string FinalPathsString { get; private set; }
        private List<string> _paths = new List<string>();
        private string _dbName, _tableName, _targetFolder;
        private Action<string> _deleteAction;
        private FlowLayoutPanel _flpList;

        public AttachmentManagerUI(string currentRelPathStr, string dbName, string tableName, string targetFolder, Action<string> deleteAction) 
        {
            _dbName = dbName; 
            _tableName = tableName; 
            _targetFolder = targetFolder; 
            _deleteAction = deleteAction;
            
            if (!string.IsNullOrEmpty(currentRelPathStr)) 
            { 
                _paths = new List<string>(currentRelPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)); 
            }
            
            this.Text = "多檔附件管理中心"; 
            this.Size = new Size(700, 600); 
            this.StartPosition = FormStartPosition.CenterParent; 
            this.FormBorderStyle = FormBorderStyle.FixedDialog; 
            this.MaximizeBox = false; 
            this.MinimizeBox = false; 
            this.BackColor = Color.White;

            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));

            GroupBox boxList = new GroupBox { Text = "1. 已上傳檔案清單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
            _flpList = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxList.Controls.Add(_flpList); 
            tlp.Controls.Add(boxList, 0, 0);

            GroupBox boxUpload = new GroupBox { Text = "2. 新增附件檔案", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
            Panel pnlDrop = new Panel { Dock = DockStyle.Fill, AllowDrop = true, BackColor = Color.AliceBlue, Cursor = Cursors.Hand };
            pnlDrop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlDrop.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Dashed);
            
            Label lblDrop = new Label { Text = "📁 點擊此處選擇多個檔案\n\n或\n\n將檔案拖曳至此區域", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.SteelBlue };
            lblDrop.Click += (s, e) => SelectFiles(); 
            pnlDrop.Click += (s, e) => SelectFiles(); 
            pnlDrop.Controls.Add(lblDrop);
            pnlDrop.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            pnlDrop.DragDrop += (s, e) => { ProcessUpload((string[])e.Data.GetData(DataFormats.FileDrop)); };
            boxUpload.Controls.Add(pnlDrop); 
            tlp.Controls.Add(boxUpload, 0, 1);

            Button btnClearAll = new Button { Text = "🗑️ 清除此筆紀錄的所有附件", Dock = DockStyle.Fill, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(3, 5, 3, 5) };
            btnClearAll.Click += (s, e) => 
            {
                if (_paths.Count == 0) return;
                if (MessageBox.Show("確定要清除所有附件嗎？\n(實體檔案將被同步永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                { 
                    foreach (var p in _paths) _deleteAction(p); 
                    _paths.Clear(); 
                    RefreshListUI(); 
                }
            };
            tlp.Controls.Add(btnClearAll, 0, 2);

            Button btnSaveClose = new Button { Text = "💾 確認變更並返回", Dock = DockStyle.Fill, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Margin = new Padding(3, 5, 3, 5) };
            btnSaveClose.Click += (s, e) => 
            { 
                FinalPathsString = string.Join("|", _paths); 
                this.DialogResult = DialogResult.OK; 
            };
            tlp.Controls.Add(btnSaveClose, 0, 3);

            this.Controls.Add(tlp); 
            RefreshListUI();
        }

        private void RefreshListUI() 
        {
            _flpList.Controls.Clear();
            if (_paths.Count == 0) 
            { 
                _flpList.Controls.Add(new Label { Text = "(尚無任何附件)", ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(10) }); 
                return; 
            }
            
            foreach (string path in _paths) 
            {
                Panel pItem = new Panel { Width = _flpList.Width - 30, Height = 40, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                Label lName = new Label { Text = Path.GetFileName(path), Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Microsoft JhengHei UI", 11F) };
                
                Button bOpen = new Button { Text = "開啟", Width = 100, Dock = DockStyle.Right, BackColor = Color.LightGray, Cursor = Cursors.Hand };
                bOpen.Click += (s, e) => 
                { 
                    try { System.Diagnostics.Process.Start(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path)); } 
                    catch (Exception ex) { MessageBox.Show("開啟失敗：" + ex.Message); } 
                };
                
                Button bDownload = new Button { Text = "下載", Width = 100, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
                bDownload.Click += (s, e) => 
                {
                    try 
                    {
                        string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
                        if (!File.Exists(sourcePath)) 
                        { 
                            MessageBox.Show("找不到原始檔案，可能已被移動或刪除。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                            return; 
                        }
                        
                        using (SaveFileDialog sfd = new SaveFileDialog()) 
                        {
                            string fileName = Path.GetFileName(path); 
                            string ext = Path.GetExtension(path);
                            sfd.FileName = fileName; 
                            sfd.Title = "另存附件檔案"; 
                            sfd.Filter = $"檔案 (*{ext})|*{ext}|所有檔案 (*.*)|*.*";
                            if (sfd.ShowDialog() == DialogResult.OK) 
                            { 
                                File.Copy(sourcePath, sfd.FileName, true); 
                                MessageBox.Show("檔案下載完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                            }
                        }
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("下載失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                };
                
                Button bDel = new Button { Text = "刪除", Width = 100, Dock = DockStyle.Right, BackColor = Color.LightCoral, ForeColor = Color.White, Cursor = Cursors.Hand };
                bDel.Click += (s, e) => 
                { 
                    if (MessageBox.Show($"確定刪除 {Path.GetFileName(path)}?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                    { 
                        _deleteAction(path); 
                        _paths.Remove(path); 
                        RefreshListUI(); 
                    } 
                };
                
                pItem.Controls.Add(lName); 
                pItem.Controls.Add(bDel); 
                pItem.Controls.Add(bDownload); 
                pItem.Controls.Add(bOpen); 
                _flpList.Controls.Add(pItem);
            }
        }

        private void SelectFiles() 
        { 
            using (OpenFileDialog ofd = new OpenFileDialog { Title = "選擇附件檔案", Multiselect = true, Filter = "所有檔案 (*.*)|*.*" }) 
            { 
                if (ofd.ShowDialog() == DialogResult.OK) ProcessUpload(ofd.FileNames); 
            } 
        }

        private void ProcessUpload(string[] sourceFiles) 
        {
            if (sourceFiles.Length == 0) return;
            using (ImageCompressionHelper compressor = new ImageCompressionHelper()) 
            {
                string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件", _dbName, _tableName, _targetFolder);
                if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
                
                foreach (string src in sourceFiles) 
                {
                    try 
                    {
                        string ext = Path.GetExtension(src); 
                        string baseName = Path.GetFileNameWithoutExtension(src); 
                        string destName = baseName + ext; 
                        string destPath = Path.Combine(destDir, destName);
                        int count = 1; 
                        
                        while (File.Exists(destPath)) 
                        { 
                            destName = $"{baseName}_{count++}{ext}"; 
                            destPath = Path.Combine(destDir, destName); 
                        }
                        
                        compressor.ProcessAndSave(src, destPath);
                        _paths.Add($"附件/{_dbName}/{_tableName}/{_targetFolder}/{destName}");
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show($"上傳檔案 {Path.GetFileName(src)} 失敗: {ex.Message}", "錯誤"); 
                    }
                }
            }
            RefreshListUI();
        }

        private class ImageCompressionHelper : IDisposable
        {
            private readonly string[] _imageExts = { ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
            private ImageCodecInfo _jpgEncoder;
            private EncoderParameters _encoderParams;

            public ImageCompressionHelper() 
            {
                _jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                _encoderParams = new EncoderParameters(1);
                _encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L); 
            }

            public void ProcessAndSave(string srcPath, string destPath) 
            {
                string ext = Path.GetExtension(srcPath).ToLower();
                if (!_imageExts.Contains(ext)) 
                { 
                    File.Copy(srcPath, destPath); 
                    return; 
                }
                
                using (Image originalImg = Image.FromFile(srcPath)) 
                {
                    int maxSide = 1024; 
                    int origWidth = originalImg.Width; 
                    int origHeight = originalImg.Height;
                    
                    if (origWidth > maxSide || origHeight > maxSide) 
                    {
                        float ratio = Math.Min((float)maxSide / origWidth, (float)maxSide / origHeight);
                        int newWidth = (int)(origWidth * ratio); 
                        int newHeight = (int)(origHeight * ratio);
                        
                        using (Bitmap resizedImg = new Bitmap(newWidth, newHeight)) 
                        {
                            using (Graphics g = Graphics.FromImage(resizedImg)) 
                            {
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic; 
                                g.SmoothingMode = SmoothingMode.HighQuality;
                                g.PixelOffsetMode = PixelOffsetMode.HighQuality; 
                                g.CompositingQuality = CompositingQuality.HighQuality;
                                g.DrawImage(originalImg, 0, 0, newWidth, newHeight);
                            }
                            
                            if ((ext == ".jpg" || ext == ".jpeg") && _jpgEncoder != null) 
                            { 
                                resizedImg.Save(destPath, _jpgEncoder, _encoderParams); 
                            } 
                            else 
                            { 
                                resizedImg.Save(destPath, originalImg.RawFormat); 
                            }
                        }
                    } 
                    else 
                    { 
                        File.Copy(srcPath, destPath); 
                    }
                }
            }

            private ImageCodecInfo GetEncoder(ImageFormat format) 
            { 
                ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders(); 
                return codecs.FirstOrDefault(codec => codec.FormatID == format.Guid); 
            }
            
            public void Dispose() 
            { 
                if (_encoderParams != null) _encoderParams.Dispose(); 
            }
        }
    }
}
