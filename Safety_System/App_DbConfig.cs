using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace Safety_System
{
    public class App_DbConfig
    {
        private TextBox _txtPath;

        public Control GetView()
        {
            // 主排版容器
            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                RowCount = 2, 
                Padding = new Padding(15, 30, 15, 10) 
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { 
                Text = "資料庫路徑設置", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F),
                AutoSize = true, 
                Padding = new Padding(15, 25, 15, 15) 
            };

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };

            flp.Controls.Add(new Label { Text = "目前資料庫位置:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });

            _txtPath = new TextBox { 
                Width = 400, 
                ReadOnly = true, 
                Text = DataManager.BasePath, 
                Margin = new Padding(3, 7, 3, 3) 
            };
            flp.Controls.Add(_txtPath);

            Button btnBrowse = new Button { 
                Text = "瀏覽...", 
                Size = new Size(80, 35), 
                BackColor = Color.LightGray, 
                Margin = new Padding(5, 2, 3, 3) 
            };
            btnBrowse.Click += BtnBrowse_Click;
            flp.Controls.Add(btnBrowse);

            Button btnSave = new Button { 
                Text = "💾 儲存設定", 
                Size = new Size(120, 35), 
                BackColor = Color.ForestGreen, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Margin = new Padding(10, 2, 3, 3) 
            };
            btnSave.Click += BtnSave_Click;
            flp.Controls.Add(btnSave);

            box.Controls.Add(flp);
            mainLayout.Controls.Add(box, 0, 0);

            // 下方說明區
            Label lblNotice = new Label {
                Text = "※ 注意：變更路徑後，系統會自動建立指定的資料夾。若要使用舊資料，請手動將原本 DB 資料夾內的 .sqlite 檔案移動至新路徑。",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                Font = new Font("Microsoft JhengHei UI", 10F)
            };
            mainLayout.Controls.Add(lblNotice, 0, 1);

            return mainLayout;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "請選擇資料庫檔案存放的資料夾";
                fbd.SelectedPath = _txtPath.Text;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    _txtPath.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtPath.Text)) return;

            // 🟢 加入密碼驗證機制
            if (VerifyPassword())
            {
                try
                {
                    DataManager.SetBasePath(_txtPath.Text);
                    MessageBox.Show("資料庫路徑已更新！\n系統已自動切換至新路徑並清除舊連線。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("儲存路徑失敗：" + ex.Message);
                }
            }
        }

        private bool VerifyPassword()
        {
            Form prompt = new Form() { 
                Width = 450, 
                Height = 240, 
                FormBorderStyle = FormBorderStyle.FixedDialog, 
                Text = "安全授權驗證", 
                StartPosition = FormStartPosition.CenterParent 
            };
            Label lbl = new Label() { Left = 30, Top = 30, Text = "變更系統核心路徑需要管理員權限，\n請輸入授權密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 14F) };
            TextBox txt = new TextBox() { Left = 30, Top = 95, Width = 370, PasswordChar = '*', Font = new Font("Microsoft JhengHei UI", 14F) };
            Button btn = new Button() { Text = "確認執行", Left = 280, Top = 145, Width = 120, Height = 40, DialogResult = DialogResult.OK, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            
            prompt.Controls.Add(lbl); 
            prompt.Controls.Add(txt); 
            prompt.Controls.Add(btn);
            prompt.AcceptButton = btn;

            return prompt.ShowDialog() == DialogResult.OK && txt.Text == "tces";
        }
    }
}
