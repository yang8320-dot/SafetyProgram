using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_DbConfig
    {
        private TextBox _txtPath;

        public Control GetView()
        {
            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, RowCount = 2, Padding = new Padding(15, 30, 15, 10) 
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            GroupBox box = new GroupBox { 
                Text = "資料庫路徑設置", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F),
                AutoSize = true, Padding = new Padding(15, 15, 15, 10) 
            };

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            flp.Controls.Add(new Label { Text = "目前資料庫位置:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });

            _txtPath = new TextBox { Width = 400, ReadOnly = true, Text = DataManager.BasePath, Margin = new Padding(3, 7, 3, 3) };
            flp.Controls.Add(_txtPath);

            Button btnBrowse = new Button { Text = "瀏覽...", Size = new Size(80, 35), BackColor = Color.LightGray };
            btnBrowse.Click += (s, e) => {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                    if (fbd.ShowDialog() == DialogResult.OK) _txtPath.Text = fbd.SelectedPath;
                }
            };
            flp.Controls.Add(btnBrowse);

            Button btnSave = new Button { 
                Text = "💾 儲存設定", Size = new Size(120, 35), BackColor = Color.ForestGreen, ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(10, 2, 3, 3) 
            };
            
            // 🟢 儲存設定加入密碼驗證
            btnSave.Click += (s, e) => {
                if (!string.IsNullOrEmpty(_txtPath.Text) && VerifyPassword()) {
                    DataManager.SetBasePath(_txtPath.Text);
                    MessageBox.Show("路徑設定已更新！");
                }
            };
            flp.Controls.Add(btnSave);

            box.Controls.Add(flp);
            mainLayout.Controls.Add(box, 0, 0);
            return mainLayout;
        }

        private bool VerifyPassword()
        {
            Form prompt = new Form() { Width = 450, Height = 240, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "安全驗證", StartPosition = FormStartPosition.CenterParent };
            TextBox txt = new TextBox() { Left = 30, Top = 95, Width = 370, PasswordChar = '*', Font = new Font("UI", 14F) };
            Button btn = new Button() { Text = "確認", Left = 280, Top = 145, Width = 120, Height = 40, DialogResult = DialogResult.OK };
            prompt.Controls.Add(new Label() { Left = 30, Top = 30, Text = "變更核心路徑，請輸入密碼：", AutoSize = true, Font = new Font("UI", 14F) });
            prompt.Controls.Add(txt); prompt.Controls.Add(btn); prompt.AcceptButton = btn;
            return prompt.ShowDialog() == DialogResult.OK && txt.Text == "tces";
        }
    }
}
