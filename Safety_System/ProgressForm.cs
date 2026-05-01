/// FILE: Safety_System/ProgressForm.cs ///
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class ProgressForm : Form
    {
        private Label _lblStatus;
        private ProgressBar _progressBar;

        public ProgressForm(string title = "處理中...")
        {
            this.Text = title;
            this.Size = new Size(450, 150);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ControlBox = false; 
            this.BackColor = Color.White;

            _lblStatus = new Label { Location = new Point(20, 20), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
            _progressBar = new ProgressBar { Location = new Point(20, 60), Size = new Size(390, 30), Style = ProgressBarStyle.Continuous };

            this.Controls.Add(_lblStatus);
            this.Controls.Add(_progressBar);
        }

        public async Task ExecuteAsync(Func<IProgress<int>, IProgress<string>, Task> work)
        {
            var progressInt = new Progress<int>(percent => {
                if (percent >= 0 && percent <= 100) {
                    _progressBar.Value = percent;
                    _progressBar.Refresh(); // 🟢 強制刷新，防止畫面假死
                }
            });

            var progressStr = new Progress<string>(text => {
                _lblStatus.Text = text;
                _lblStatus.Refresh(); // 🟢 強制刷新
            });

            this.Shown += async (s, e) => {
                try { await Task.Run(() => work(progressInt, progressStr)); }
                catch (Exception ex) { MessageBox.Show($"發生錯誤：\n{ex.Message}"); }
                finally { this.DialogResult = DialogResult.OK; this.Close(); }
            };

            this.ShowDialog(Form.ActiveForm);
        }
    }
}
