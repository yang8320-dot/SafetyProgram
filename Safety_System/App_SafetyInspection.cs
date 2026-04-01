using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_SafetyInspection
    {
        private TextBox _txtLoc = new TextBox();
        private TextBox _txtUser = new TextBox();
        private ComboBox _cmbStatus = new ComboBox();

        public Control GetView()
        {
            var pnl = new Panel();
            
            var lblTitle = new Label { 
                Text = "工安巡檢登錄", 
                Font = new Font("Microsoft JhengHei", 18, FontStyle.Bold), 
                Location = new Point(10, 10), 
                AutoSize = true 
            };
            pnl.Controls.Add(lblTitle);

            AddField(pnl, "巡檢地點:", _txtLoc, 70);
            AddField(pnl, "巡檢人員:", _txtUser, 110);
            
            var lblStatus = new Label { Text = "狀態:", Location = new Point(10, 150), AutoSize = true };
            _cmbStatus.Location = new Point(110, 150);
            _cmbStatus.Width = 200;
            _cmbStatus.Items.AddRange(new string[] { "正常", "異常", "待處理" });
            _cmbStatus.SelectedIndex = 0;
            pnl.Controls.Add(lblStatus);
            pnl.Controls.Add(_cmbStatus);

            var btn = new Button { 
                Text = "儲存紀錄", 
                Location = new Point(110, 200), 
                Size = new Size(100, 40), 
                BackColor = Color.LightSteelBlue 
            };
            btn.Click += (s, e) => {
                DataManager.SaveInspectionRecord(DateTime.Now.ToString("yyyy/MM/dd HH:mm"), _txtLoc.Text, _txtUser.Text, _cmbStatus.Text);
                MessageBox.Show("紀錄儲存成功！");
                _txtLoc.Clear();
                _txtUser.Clear();
            };
            pnl.Controls.Add(btn);

            return pnl;
        }

        private void AddField(Panel p, string label, Control input, int y)
        {
            p.Controls.Add(new Label { Text = label, Location = new Point(10, y), AutoSize = true });
            input.Location = new Point(110, y);
            input.Width = 200;
            p.Controls.Add(input);
        }
    }
}
