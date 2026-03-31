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

        Button btnSave = new Button() { Text = "儲存檔案", Left = 110, Top = 5, Width = 100, Height = 35, BackColor = Color.Gainsboro, FlatStyle = FlatStyle.Flat, Font = MainFont, Cursor = Cursors.Hand };
        btnSave.Click += BtnSave_Click;

        statusLabel = new Label() { Text = "點擊「新增截圖」開始...", Left = 220, Top = 12, AutoSize = true, Font = MainFont, ForeColor = Color.DimGray };

        topPanel.Controls.AddRange(new Control[] { btnNew, btnSave, statusLabel });
        this.Controls.Add(topPanel);

        // --- 預覽區 ---
        Panel previewContainer = new Panel() { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(5) };
        previewBox = new PictureBox() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(240, 240, 240) };
        previewContainer.Controls.Add(previewBox);
        
        this.Controls.Add(previewContainer);
        previewContainer.BringToFront();
    }

    private void BtnNew_Click(object sender, EventArgs e) {
        // 1. 隱藏主視窗，避免擋住畫面
        parentForm.Hide();
        
        // 2. 暫停一下，確保主視窗完全消失的動畫跑完，再進行螢幕截取
        Application.DoEvents();
        Thread.Sleep(250); 

        // 3. 開啟透明截圖覆蓋層
        using (SnippingOverlayForm snipForm = new SnippingOverlayForm()) {
            if (snipForm.ShowDialog() == DialogResult.OK && snipForm.ResultImage != null) {
                // 4. 將結果顯示在預覽框
                if (previewBox.Image != null) previewBox.Image.Dispose();
                previewBox.Image = (Image)snipForm.ResultImage.Clone();
                
                // 5. 自動放入系統剪貼簿
                Clipboard.SetImage(previewBox.Image);
                statusLabel.Text = "截圖成功！已複製到剪貼簿。";
                statusLabel.ForeColor = Color.FromArgb(0, 153, 76);
            } else {
                statusLabel.Text = "已取消截圖。";
                statusLabel.ForeColor = Color.DimGray;
            }
        }

        // 6. 截圖結束，恢復顯示主視窗
        parentForm.ShowAppWindow(4); // 假設截圖是第 5 個分頁 (Index 4)
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
            }
        }
    }
}

// ==========================================
// 核心：全螢幕截圖覆蓋視窗 (Overlay)
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
        this.DoubleBuffered = true; // 防止畫面閃爍

        // 設定範圍涵蓋所有螢幕 (支援雙螢幕)
        this.Location = SystemInformation.VirtualScreen.Location;
        this.Size = SystemInformation.VirtualScreen.Size;

        // 瞬間捕捉全螢幕畫面作為底圖
        screenBmp = new Bitmap(this.Width, this.Height);
        using (Graphics g = Graphics.FromImage(screenBmp)) {
            g.CopyFromScreen(this.Location, Point.Empty, this.Size);
        }
    }

    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        
        // 1. 畫上捕捉到的原圖
        e.Graphics.DrawImageUnscaled(screenBmp, 0, 0);

        // 2. 畫一層半透明黑色，讓螢幕變暗 (模擬截圖工具效果)
        using (SolidBrush dimBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0))) {
            e.Graphics.FillRectangle(dimBrush, this.ClientRectangle);
        }

        // 3. 如果正在拖曳，就把選取範圍「洗亮」(畫回原圖)，並加上紅框
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
            // 右鍵取消截圖
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
            this.Invalidate(); // 要求重繪畫面
        }
    }

    protected override void OnMouseUp(MouseEventArgs e) {
        if (isDragging) {
            isDragging = false;
            // 確保有框選到足夠的範圍 (防呆，避免誤點)
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
