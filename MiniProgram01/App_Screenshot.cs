using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Threading;

public class App_Screenshot : UserControl {
    private MainForm parentForm;
    private PictureBox previewBox;
    private Label statusLabel;
    
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);

    public App_Screenshot(MainForm mainForm) {
        this.parentForm = mainForm;
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        // --- 頂部控制列 ---
        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = 50 };
        
        Button btnNew = new Button() { Text = "新增截圖", Left = 0, Top = 5, Width = 100, Height = 35, BackColor = AppleBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font(MainFont, FontStyle.Bold), Cursor = Cursors.Hand };
        btnNew.Click += BtnNew_Click;

        // 【新增】複製圖片按鈕
        Button btnCopy = new Button() { Text = "複製圖片", Left = 110, Top = 5, Width = 100, Height = 35, BackColor = Color.FromArgb(0, 153, 76), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = MainFont, Cursor = Cursors.Hand };
        btnCopy.Click += BtnCopy_Click;

        // 【調整位置】將儲存按鈕往右移
        Button btnSave = new Button() { Text = "儲存檔案", Left = 220, Top = 5, Width = 100, Height = 35, BackColor = Color.Gainsboro, FlatStyle = FlatStyle.Flat, Font = MainFont, Cursor = Cursors.Hand };
        btnSave.Click += BtnSave_Click;

        // 【調整位置】將狀態文字往右移
        statusLabel = new Label() { Text = "點擊「新增截圖」開始...", Left = 330, Top = 12, AutoSize = true, Font = MainFont, ForeColor = Color.DimGray };

        topPanel.Controls.AddRange(new Control[] { btnNew, btnCopy, btnSave, statusLabel });
        this.Controls.Add(topPanel);

        // --- 預覽區 ---
        Panel previewContainer = new Panel() { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(5) };
        previewBox = new PictureBox() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(240, 240, 240) };
        previewContainer.Controls.Add(previewBox);
        
        this.Controls.Add(previewContainer);
        previewContainer.BringToFront();
    }

    private void BtnNew_Click(object sender, EventArgs e) {
        parentForm.Hide();
        Application.DoEvents();
        Thread.Sleep(250); 

        using (SnippingOverlayForm snipForm = new SnippingOverlayForm()) {
            if (snipForm.ShowDialog() == DialogResult.OK && snipForm.ResultImage != null) {
                if (previewBox.Image != null) previewBox.Image.Dispose();
                previewBox.Image = (Image)snipForm.ResultImage.Clone();
                
                // 截圖當下依然保持自動複製功能，體驗更順暢
                Clipboard.SetImage(previewBox.Image);
                statusLabel.Text = "截圖成功！已自動複製。";
                statusLabel.ForeColor = Color.FromArgb(0, 153, 76);
            } else {
                statusLabel.Text = "已取消截圖。";
                statusLabel.ForeColor = Color.DimGray;
            }
        }

        parentForm.ShowAppWindow(4); 
    }

    // 【新增】手動點擊複製的觸發事件
    private void BtnCopy_Click(object sender, EventArgs e) {
        if (previewBox.Image == null) {
            MessageBox.Show("目前沒有可複製的截圖！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        
        Clipboard.SetImage(previewBox.Image);
        statusLabel.Text = "已重新複製到剪貼簿！";
        statusLabel.ForeColor = Color.FromArgb(0, 122, 255); // 改成藍色字體給予視覺回饋
    }

    private void BtnSave_Click(object sender, EventArgs e) {
        if (previewBox.Image == null) {
            MessageBox.Show("目前沒有可儲存的截圖！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using (SaveFileDialog sfd = new SaveFileDialog()) {
            sfd.Filter = "PNG 圖片|*.png|JPEG 圖片|*.jpg";
            sfd.FileName = "截圖_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            if (sfd.ShowDialog() == DialogResult.OK) {
                ImageFormat format = sfd.FileName.EndsWith(".jpg") ? ImageFormat.Jpeg : ImageFormat.Png;
                previewBox.Image.Save(sfd.FileName, format);
                statusLabel.Text = "檔案已成功儲存！";
                statusLabel.ForeColor = Color.DimGray;
            }
        }
    }
}

// ==========================================
// 核心：全螢幕截圖覆蓋視窗 (Overlay) (保持不變)
// ==========================================
public class SnippingOverlayForm : Form {
    private Bitmap screenBmp;
    private Point startPt;
    private Rectangle selectRect;
    private bool isDragging = false;
    
    public Image ResultImage { get; private set; }

    public SnippingOverlayForm() {
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.Cursor = Cursors.Cross;
        this.DoubleBuffered = true; 

        this.Location = SystemInformation.VirtualScreen.Location;
        this.Size = SystemInformation.VirtualScreen.Size;

        screenBmp = new Bitmap(this.Width, this.Height);
        using (Graphics g = Graphics.FromImage(screenBmp)) {
            g.CopyFromScreen(this.Location, Point.Empty, this.Size);
        }
    }

    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        
        e.Graphics.DrawImageUnscaled(screenBmp, 0, 0);

        using (SolidBrush dimBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0))) {
            e.Graphics.FillRectangle(dimBrush, this.ClientRectangle);
        }

        if (selectRect.Width > 0 && selectRect.Height > 0) {
            e.Graphics.DrawImage(screenBmp, selectRect, selectRect, GraphicsUnit.Pixel);
            using (Pen borderPen = new Pen(Color.Red, 2)) {
                e.Graphics.DrawRectangle(borderPen, selectRect);
            }
        }
    }

    protected override void OnMouseDown(MouseEventArgs e) {
        if (e.Button == MouseButtons.Left) {
            isDragging = true;
            startPt = e.Location;
            selectRect = new Rectangle(e.X, e.Y, 0, 0);
        } 
        else if (e.Button == MouseButtons.Right) {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    protected override void OnMouseMove(MouseEventArgs e) {
        if (isDragging) {
            int x = Math.Min(startPt.X, e.X);
            int y = Math.Min(startPt.Y, e.Y);
            int w = Math.Abs(startPt.X - e.X);
            int h = Math.Abs(startPt.Y - e.Y);
            selectRect = new Rectangle(x, y, w, h);
            this.Invalidate(); 
        }
    }

    protected override void OnMouseUp(MouseEventArgs e) {
        if (isDragging) {
            isDragging = false;
            if (selectRect.Width > 5 && selectRect.Height > 5) {
                ResultImage = new Bitmap(selectRect.Width, selectRect.Height);
                using (Graphics g = Graphics.FromImage(ResultImage)) {
                    g.DrawImage(screenBmp, new Rectangle(0, 0, selectRect.Width, selectRect.Height), selectRect, GraphicsUnit.Pixel);
                }
                this.DialogResult = DialogResult.OK;
            } else {
                this.DialogResult = DialogResult.Cancel;
            }
            this.Close();
        }
    }

    protected override void Dispose(bool disposing) {
        if (disposing && screenBmp != null) {
            screenBmp.Dispose();
        }
        base.Dispose(disposing);
    }
}
