using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace Safety_System
{
    public class App_DbConfig : Form
    {
        private TextBox _txtPath = new TextBox();

        public App_DbConfig()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "資料庫路徑設定";
            this.Size = new Size(600, 250);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            // 標題文字串
            Label lblHeader = new Label
            {
                Text = "數據資料存放設定",
                Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold),
                Location = new Point(20, 20),
                AutoSize = true
            };
            this.Controls.Add(lblHeader);

            // 路徑顯示欄位
            _txtPath.Location = new Point(20, 70);
            _txtPath.Width = 400;
            _txtPath.ReadOnly = true;
            _txtPath.Text = DataManager.BasePath; // 顯示目前路徑
            this.Controls.Add(_txtPath);

            // 選擇資料夾按鍵
            Button btnBrowse = new Button
            {
                Text = "選擇資料夾",
                Location = new Point(430, 68),
                Size = new Size(120, 35),
                BackColor = Color.LightGray
            };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            // 確認儲存按鍵 (放在下兩行位置)
            Button btnSave = new Button
            {
                Text = "確認儲存設定",
                Location = new Point(20, 130),
                Size = new Size(530, 45),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "請選擇數據資料存放的資料夾";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    _txtPath.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(_txtPath.Text))
            {
                DataManager.SetBasePath(_txtPath.Text);
                MessageBox.Show("資料存放路徑已成功更新！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close(); // 儲存後自動關閉
            }
            else
            {
                MessageBox.Show("請選擇有效的資料夾路徑。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
