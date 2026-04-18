/*
 * 檔案功能：全域畫面截圖模組 (支援自訂範圍拖曳截取、雙螢幕捕捉、複製至剪貼簿與存檔)
 * 對應選單名稱：畫面截圖
 * 對應資料庫名稱：無 (直接存取圖片檔與剪貼簿)
 * 資料表名稱：無
 */

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

public class App_Screenshot : UserControl
{
    private MainForm parentForm;
    private PictureBox pictureBox;
    private Label lblStatus;
    
    // --- 樣式設定 (iOS 風格) ---
    private static Color AppleBgColor = Color.FromArgb(245, 245, 247);
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Color AppleRed = Color.FromArgb(255, 59, 48);
    private static Color AppleGreen = Color.FromArgb(52, 199, 89);
    private static Color AppleGray = Color.FromArgb(142, 142, 147);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Regular);
    private static Font BoldFont = new Font("Microsoft JhengHei UI", 11f, FontStyle.Bold);

    public App_Screenshot(MainForm mainForm)
    {
        this.parentForm = mainForm;
        
        // 1. 初始化控制項與 DPI 支援
        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.BackColor = AppleBgColor;
        this.Font = MainFont;
        this.Padding = new Padding(15); // 與主框架間保持 15 的留白
        
        // 2. 建構純程式碼 UI
        InitializeUI();
    }

    /// <summary>
    /// 建構 iOS 風格純程式碼介面 (Code-First UI)
    /// </summary>
    private void InitializeUI()
    {
        // ==========================================
        // 頂部控制區塊 (Action Panel)
        // ==========================================
        TableLayoutPanel headerPanel = new TableLayoutPanel()
        {
            Dock = DockStyle.Top,
            Height = 45,
            ColumnCount = 5,
            BackColor = Color.Transparent,
            Margin = new Padding(0, 0, 0, 15)
        };
        
        // 分配欄寬 (四個按鈕 + 一個狀態標籤)
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        Button btnCapture = CreateAppleButton("截圖", AppleBlue);
        btnCapture.Click += BtnCapture_Click;

        Button btnCopy = CreateAppleButton("複製", AppleGreen);
        btnCopy.Click += BtnCopy_Click;

        Button btnSave = CreateAppleButton("儲存", AppleGray);
        btnSave.Click += BtnSave_Click;

        Button btnClear = CreateAppleButton("清除", AppleRed);
        btnClear.Click += BtnClear_Click;

        lblStatus = new Label()
        {
            Text = "等待操作...",
            Font = MainFont,
            ForeColor = Color.Gray,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleRight,
            Padding = new Padding(0, 0, 10, 0) // 狀態文字離右邊界留白 10
        };

        headerPanel.Controls.Add(btnCapture, 0, 0);
        headerPanel.Controls.Add(btnCopy, 1, 0);
        headerPanel.Controls.Add(btnSave, 2, 0);
        headerPanel.Controls.Add(btnClear, 3, 0);
        headerPanel.Controls.Add(lblStatus, 4, 0);

        // ==========================================
        // 中間圖片顯示區塊 (iOS Style Card)
        // ==========================================
        Panel imageCard = new Panel()
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(10) // 圖片與外框保留 10 的內縮距離
        };

        pictureBox = new PictureBox()
        {
            Dock = DockStyle.Fill,
            SizeMode = PictureBoxSizeMode.Zoom, // 保持圖片比例自動縮放
            BackColor = Color.FromArgb(250, 250, 250), // 圖片背景微灰，與白底產生對比
        };

        imageCard.Controls.Add(pictureBox);

        // ==========================================
        // 組合至主畫面
        // ==========================================
        this.Controls.Add(imageCard);
        this.Controls.Add(new Panel() { Dock = DockStyle.Top, Height = 15, BackColor = Color.Transparent }); // 確保間隔為 15
        this.Controls.Add(headerPanel);
    }

    /// <summary>
    /// 動態生成扁平化 iOS 風格按鈕
    /// </summary>
    private Button CreateAppleButton(string text, Color bgColor)
    {
        Button btn = new Button()
        {
            Text = text,
            Dock = DockStyle.Fill,
            BackColor = bgColor,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = BoldFont,
            Margin = new Padding(0, 5, 10, 5) // 按鈕之間保留間隔
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    // ==========================================
    // 截圖與影像處理邏輯
    // ==========================================

    private void BtnCapture_Click(object sender, EventArgs e)
    {
        // 隱藏主視窗，避免被截圖擷取到
        parentForm.Hide();
        System.Threading.Thread.Sleep(300); // 確保動畫完全隱藏視窗

        // 啟動全螢幕截圖遮罩
        using (SnippingOverlayForm overlay = new SnippingOverlayForm())
        {
            if (overlay.ShowDialog() == DialogResult.OK && overlay.ResultImage != null)
            {
                // 釋放舊圖片記憶體
                if (pictureBox.Image != null) pictureBox.Image.Dispose();
                
                pictureBox.Image = new Bitmap(overlay.ResultImage);
                Clipboard.SetImage(pictureBox.Image);
                
                lblStatus.Text = "截圖完成並已自動複製至剪貼簿！";
                lblStatus.ForeColor = AppleGreen;
            }
            else
            {
                lblStatus.Text = "截圖已取消。";
                lblStatus.ForeColor = Color.Gray;
            }
        }

        // 截圖結束後自動恢復主視窗並跳轉回此分頁
        parentForm.ShowAppWindow();
    }

    private void BtnCopy_Click(object sender, EventArgs e)
    {
        if (pictureBox.Image != null)
        {
            Clipboard.SetImage(pictureBox.Image);
            lblStatus.Text = "已複製到剪貼簿！";
            lblStatus.ForeColor = AppleBlue;
        }
        else
        {
            lblStatus.Text = "請先進行截圖。";
            lblStatus.ForeColor = AppleRed;
        }
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        if (pictureBox.Image == null)
        {
            lblStatus.Text = "沒有可儲存的圖片。";
            lblStatus.ForeColor = AppleRed;
            return;
        }

        using (SaveFileDialog sfd = new SaveFileDialog())
        {
            sfd.Filter = "PNG 圖片|*.png|JPEG 圖片|*.jpg|BMP 圖片|*.bmp";
            sfd.Title = "儲存截圖";
            sfd.FileName = $"Screenshot_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ImageFormat format = ImageFormat.Png;
                if (sfd.FileName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)) format = ImageFormat.Jpeg;
                else if (sfd.FileName.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)) format = ImageFormat.Bmp;

                pictureBox.Image.Save(sfd.FileName, format);
                lblStatus.Text = "儲存成功！";
                lblStatus.ForeColor = AppleBlue;
            }
        }
    }

    private void BtnClear_Click(object sender, EventArgs e)
    {
        if (pictureBox.Image != null)
        {
            pictureBox.Image.Dispose();
            pictureBox.Image = null;
        }
        lblStatus.Text = "圖片已清除。";
        lblStatus.ForeColor = Color.Gray;
    }
}

// ==========================================
// 視窗：全螢幕半透明遮罩 (負責滑鼠拖曳與影像捕捉)
// ==========================================
public class SnippingOverlayForm : Form
{
    private Bitmap screenBmp;
    private Point startPt;
    private Rectangle selectRect;
    private bool isDragging = false;
    
    public Image ResultImage { get; private set; }

    public SnippingOverlayForm()
    {
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.Cursor = Cursors.Cross;
        this.DoubleBuffered = true; // 啟用雙重緩衝，防止拖曳時畫面閃爍

        // 取得全域虛擬螢幕大小 (完整支援多螢幕)
        Rectangle bounds = SystemInformation.VirtualScreen;
        this.Bounds = bounds;

        // 捕捉當前全螢幕畫面
        screenBmp = new Bitmap(bounds.Width, bounds.Height);
        using (Graphics g = Graphics.FromImage(screenBmp))
        {
            g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);
        }

        // 將截取的靜態畫面作為背景
        this.BackgroundImage = screenBmp;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        // 1. 繪製全畫面半透明黑色遮罩 (ARGB: 120 透明度)
        using (SolidBrush dimBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0)))
        {
            e.Graphics.FillRectangle(dimBrush, this.ClientRectangle);
        }

        // 2. 將滑鼠框選的區域「挖空」露出底下明亮的截圖
        if (selectRect.Width > 0 && selectRect.Height > 0)
        {
            e.Graphics.DrawImage(screenBmp, selectRect, selectRect, GraphicsUnit.Pixel);
            
            // 3. 繪製精緻的藍色選取外框 (iOS Blue)
            using (Pen borderPen = new Pen(Color.FromArgb(0, 122, 255), 2))
            {
                e.Graphics.DrawRectangle(borderPen, selectRect);
            }
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            isDragging = true;
            startPt = e.Location;
            selectRect = new Rectangle(e.X, e.Y, 0, 0);
        }
        else if (e.Button == MouseButtons.Right)
        {
            // 按下滑鼠右鍵可直接取消截圖作業
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (isDragging)
        {
            // 動態計算滑鼠拖曳時的矩形寬高 (支援向任意方向拖曳)
            int x = Math.Min(startPt.X, e.X);
            int y = Math.Min(startPt.Y, e.Y);
            int w = Math.Abs(startPt.X - e.X);
            int h = Math.Abs(startPt.Y - e.Y);
            
            selectRect = new Rectangle(x, y, w, h);
            this.Invalidate(); // 觸發 OnPaint 重新繪製畫面
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (isDragging)
        {
            isDragging = false;
            
            // 避免使用者只是單點滑鼠而沒有拖曳出範圍
            if (selectRect.Width > 5 && selectRect.Height > 5)
            {
                ResultImage = new Bitmap(selectRect.Width, selectRect.Height);
                using (Graphics g = Graphics.FromImage(ResultImage))
                {
                    g.DrawImage(screenBmp, new Rectangle(0, 0, selectRect.Width, selectRect.Height), selectRect, GraphicsUnit.Pixel);
                }
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                this.DialogResult = DialogResult.Cancel;
            }
            this.Close();
        }
    }

    // 確實釋放全螢幕 Bitmap 記憶體
    protected override void Dispose(bool disposing)
    {
        if (disposing && screenBmp != null)
        {
            screenBmp.Dispose();
        }
        base.Dispose(disposing);
    }
}
