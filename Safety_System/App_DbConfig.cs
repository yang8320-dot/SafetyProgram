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
                AutoSize = true, Padding = new Padding(15, 15, 15, 5) 
            };

            FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            flp.Controls.Add(new Label { Text = "目前位置:", AutoSize = true, Margin = new Padding(3, 10, 3, 0) });

            _txtPath = new TextBox { Width = 400, ReadOnly = true, Text = DataManager.BasePath, Margin = new Padding(3, 7, 3, 3) };
            flp.Controls.Add(_txtPath);

            Button btnBrowse = new Button { Text = "瀏覽...", Size = new Size(80, 35), BackColor = Color.LightGray };
            btnBrowse.Click += (s, e) => {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                    if (fbd.ShowDialog() == DialogResult.OK) _txtPath.Text = fbd.SelectedPath;
                }
            };
            flp.Controls.Add(btnBrowse);

            Button btnSave = new Button { Text = "儲存", Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White };
            btnSave.Click += (s, e) => {
                if (VerifyPassword()) {
                    DataManager.SetBasePath(_txtPath.Text);
                    MessageBox.Show("設定已更新");
                }
            };
            flp.Controls.Add(btnSave);

            box.Controls.Add(flp);
            mainLayout.Controls.Add(box, 0, 0);
            return mainLayout;
        }

        private bool VerifyPassword()
        {
            Form prompt = new Form() { Width = 400, Height = 200, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "安全驗證", StartPosition = FormStartPosition.CenterParent };
            TextBox txt = new TextBox() { Left = 50, Top = 50, Width = 280, PasswordChar = '*' };
            Button btn = new Button() { Text = "確認", Left = 250, Top = 100, DialogResult = DialogResult.OK };
            prompt.Controls.Add(new Label() { Left = 50, Top = 20, Text = "請輸入密碼 (tces):" });
            prompt.Controls.Add(txt); prompt.Controls.Add(btn); prompt.AcceptButton = btn;
            return prompt.ShowDialog() == DialogResult.OK && txt.Text == "tces";
        }
    }
}
