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
        
        // 🟢 儲存每個項目的 CheckBox，用來判斷哪些被勾選
        private Dictionary<string, CheckBox> _checkBoxMap = new Dictionary<string, CheckBox>();

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
            this.Size = new Size(750, 680); // 🟢 視窗加寬 50 (原700 -> 750)
            this.StartPosition = FormStartPosition.CenterParent; 
            this.FormBorderStyle = FormBorderStyle.FixedDialog; 
            this.MaximizeBox = false; 
            this.MinimizeBox = false; 
            this.BackColor = Color.White;

            TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5 };
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));

            GroupBox boxList = new GroupBox { Text = "1. 已上傳檔案清單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
            _flpList = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxList.Controls.Add(_flpList); 
            tlp.Controls.Add(boxList, 0, 0);

            // 🟢 批量操作列
            Panel pnlBatch = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            Button btnSelectAll = new Button { Text = "☑️ 全選", Width = 100, Dock = DockStyle.Left, BackColor = Color.LightGray, Cursor = Cursors.Hand };
            btnSelectAll.Click += (s, e) => { foreach (var cb in _checkBoxMap.Values) cb.Checked = true; };
            
            Button btnUnselectAll = new Button { Text = "☐ 取消全選", Width = 110, Dock = DockStyle.Left, BackColor = Color.LightGray, Cursor = Cursors.Hand };
            btnUnselectAll.Click += (s, e) => { foreach (var cb in _checkBoxMap.Values) cb.Checked = false; };
            
            Button btnBatchExport = new Button { Text = "💾 批量轉存勾選檔案", Width = 200, Dock = DockStyle.Right, BackColor = Color.DarkCyan, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnBatchExport.Click += BtnBatchExport_Click;

            pnlBatch.Controls.Add(btnUnselectAll);
            pnlBatch.Controls.Add(btnSelectAll);
            pnlBatch.Controls.Add(btnBatchExport);
            tlp.Controls.Add(pnlBatch, 0, 1);

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
            tlp.Controls.Add(boxUpload, 0, 2);

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
            tlp.Controls.Add(btnClearAll, 0, 3);

            Button btnSaveClose = new Button { Text = "✔️ 確認變更並返回", Dock = DockStyle.Fill, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Margin = new Padding(3, 5, 3, 5) };
            btnSaveClose.Click += (s, e) => 
            { 
                FinalPathsString = string.Join("|", _paths); 
                this.DialogResult = DialogResult.OK; 
            };
            tlp.Controls.Add(btnSaveClose, 0, 4);

            this.Controls.Add(tlp); 
            RefreshListUI();
        }

        private void RefreshListUI() 
        {
            _flpList.Controls.Clear();
            _checkBoxMap.Clear();

            if (_paths.Count == 0) 
            { 
                _flpList.Controls.Add(new Label { Text = "(尚無任何附件)", ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(10) }); 
                return; 
            }
            
            foreach (string path in _paths) 
            {
                Panel pItem = new Panel { Width = _flpList.Width - 30, Height = 40, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                
                CheckBox cb = new CheckBox { Dock = DockStyle.Left, Width = 30, Padding = new Padding(5,0,0,0) };
                _checkBoxMap[path] = cb; 

                Label lName = new Label { Text = Path.GetFileName(path), Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Microsoft JhengHei UI", 11F) };
                
                // 🟢 調整：按鈕寬度縮小到 70 搭配加寬的介面放入新按鈕
                Button bOpen = new Button { Text = "開啟", Width = 70, Dock = DockStyle.Right, BackColor = Color.LightGray, Cursor = Cursors.Hand };
                bOpen.Click += (s, e) => 
                { 
                    try 
                    { 
                        string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
                        if (File.Exists(sourcePath))
                        {
                            string tempFolder = Path.Combine(Path.GetTempPath(), "SafetySystem_Viewer");
                            if (!Directory.Exists(tempFolder)) Directory.CreateDirectory(tempFolder);
                            
                            string tempFileName = $"{DateTime.Now.ToString("HHmmss")}_{Path.GetFileName(path)}";
                            string tempFilePath = Path.Combine(tempFolder, tempFileName);
                            
                            File.Copy(sourcePath, tempFilePath, true);
                            System.Diagnostics.Process.Start(tempFilePath); 
                        }
                        else
                        {
                            MessageBox.Show("找不到原始實體檔案，可能已被其他使用者移動或刪除。", "開啟失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("開啟失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    } 
                };
                
                // 🟢 更名：將「單檔下載」改為「下載」
                Button bDownload = new Button { Text = "下載", Width = 70, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
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

                // 🟢 新增：獨立更名按鈕
                Button bRename = new Button { Text = "更名", Width = 70, Dock = DockStyle.Right, BackColor = Color.DarkOrange, ForeColor = Color.White, Cursor = Cursors.Hand };
                bRename.Click += (s, e) => 
                {
                    string oldFileName = Path.GetFileName(path);
                    string newFileName = ShowInputBox("請輸入新的檔案名稱 (包含副檔名)：", "附件更名", oldFileName);

                    if (string.IsNullOrWhiteSpace(newFileName) || newFileName == oldFileName) return;

                    // 檢查是否含有不合法字元
                    if (newFileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
                        MessageBox.Show("檔名包含無效字元！請勿輸入 \\ / : * ? \" < > |", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    try {
                        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                        string oldAbsPath = Path.Combine(baseDir, path);
                        
                        // 取得原本所在的相對目錄路徑 (結尾加上斜線)
                        string directoryRelPath = path.Substring(0, path.LastIndexOf('/') + 1);
                        string newRelPath = directoryRelPath + newFileName;
                        string newAbsPath = Path.Combine(baseDir, newRelPath);

                        if (File.Exists(newAbsPath) && oldAbsPath.ToLower() != newAbsPath.ToLower()) {
                            MessageBox.Show("該資料夾下已存在同名檔案，請更換名稱！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        // 變更實體檔案名稱
                        if (File.Exists(oldAbsPath)) {
                            File.Move(oldAbsPath, newAbsPath);
                        } else {
                            MessageBox.Show("找不到實體檔案，可能已被刪除！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // 更新記憶體中的路徑清單
                        int idx = _paths.IndexOf(path);
                        if (idx >= 0) _paths[idx] = newRelPath;

                        // 重新繪製 UI
                        RefreshListUI();
                    } catch (Exception ex) {
                        MessageBox.Show("更名失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };
                
                Button bDel = new Button { Text = "刪除", Width = 70, Dock = DockStyle.Right, BackColor = Color.LightCoral, ForeColor = Color.White, Cursor = Cursors.Hand };
                bDel.Click += (s, e) => 
                { 
                    if (MessageBox.Show($"確定刪除 {Path.GetFileName(path)}?\n(這將從硬碟永久移除檔案)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                    { 
                        _deleteAction(path); 
                        _paths.Remove(path); 
                        RefreshListUI(); 
                    } 
                };
                
                pItem.Controls.Add(lName);
                pItem.Controls.Add(cb); 
                pItem.Controls.Add(bDel); 
                pItem.Controls.Add(bRename); // 🟢 加入更名按鈕
                pItem.Controls.Add(bDownload); 
                pItem.Controls.Add(bOpen); 
                _flpList.Controls.Add(pItem);
            }
        }

        // 🟢 輔助方法：彈出輸入框供使用者修改檔名
        private string ShowInputBox(string prompt, string title, string defaultValue)
        {
            using (Form form = new Form())
            {
                form.Width = 450;
                form.Height = 200;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.BackColor = Color.White;

                Label label = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox textBox = new TextBox() { Left = 20, Top = 55, Width = 390, Text = defaultValue, Font = new Font("Microsoft JhengHei UI", 12F) };

                Button confirmation = new Button() { Text = "確認", Left = 200, Width = 100, Height = 40, Top = 100, DialogResult = DialogResult.OK, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                Button cancel = new Button() { Text = "取消", Left = 310, Width = 100, Height = 40, Top = 100, DialogResult = DialogResult.Cancel, Font = new Font("Microsoft JhengHei UI", 11F) };

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(confirmation);
                form.Controls.Add(cancel);
                form.AcceptButton = confirmation;

                return form.ShowDialog(this) == DialogResult.OK ? textBox.Text : "";
            }
        }

        // 實作批量轉存邏輯
        private void BtnBatchExport_Click(object sender, EventArgs e)
        {
            var selectedPaths = _checkBoxMap.Where(kvp => kvp.Value.Checked).Select(kvp => kvp.Key).ToList();

            if (selectedPaths.Count == 0)
            {
                MessageBox.Show("請先勾選您想要轉存的檔案！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "請選擇要把這些附件存放到哪一個資料夾：" })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    string targetDir = fbd.SelectedPath;
                    int successCount = 0;
                    int failCount = 0;

                    foreach (string relPath in selectedPaths)
                    {
                        try
                        {
                            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relPath);
                            if (File.Exists(sourcePath))
                            {
                                string fileName = Path.GetFileName(relPath);
                                string destPath = Path.Combine(targetDir, fileName);

                                // 若遇到同名檔案，自動加上數字後綴避免覆蓋
                                int count = 1;
                                string baseName = Path.GetFileNameWithoutExtension(fileName);
                                string ext = Path.GetExtension(fileName);

                                while (File.Exists(destPath))
                                {
                                    destPath = Path.Combine(targetDir, $"{baseName}_{count++}{ext}");
                                }

                                File.Copy(sourcePath, destPath, false);
                                successCount++;
                            }
                            else
                            {
                                failCount++;
                            }
                        }
                        catch
                        {
                            failCount++;
                        }
                    }

                    if (failCount == 0)
                    {
                        MessageBox.Show($"成功將 {successCount} 個檔案轉存至目標資料夾！", "批量轉存完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"完成！成功 {successCount} 個，失敗 {failCount} 個。\n(失敗原因可能為原檔案遺失或沒有權限寫入該資料夾)", "批量轉存報告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
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
