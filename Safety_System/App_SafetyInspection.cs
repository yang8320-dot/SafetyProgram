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
            Panel pnl = new Panel();
            Label lblTitle = new Label { 
                Text = "工安巡檢登錄 (SQLite 版本)", 
                Font = new Font("Microsoft JhengHei UI", 20, FontStyle.Bold), 
                Location = new Point(20, 20), 
                AutoSize = true 
            };
            pnl.Controls.Add(lblTitle);

            AddRow(pnl, "巡檢地點:", _txtLoc, 100);
            AddRow(pnl, "巡檢人員:", _txtUser, 160);
            
            Label lblStatus = new Label { 
                Text = "設備狀態:", 
                Font = new Font("Microsoft JhengHei UI", 14), 
                Location = new Point(20, 220), 
                AutoSize = true 
            };
            _cmbStatus.Location = new Point(180, 220);
            _cmbStatus.Width = 350;
            _cmbStatus.Font = new Font("Microsoft JhengHei UI", 14);
            _cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbStatus.Items.AddRange(new object[] { "正常", "異常", "待派修" });
            _cmbStatus.SelectedIndex = 0;
            pnl.Controls.Add(lblStatus);
            pnl.Controls.Add(_cmbStatus);

            Button btn = new Button { 
                Text = "儲存至資料庫", 
                Location = new Point(180, 300), 
                Size = new Size(200, 60), 
                BackColor = Color.LightSteelBlue,
                Font = new Font("Microsoft JhengHei UI", 14, FontStyle.Bold)
            };
            btn.Click += (s, e) => {
                if (string.IsNullOrEmpty(_txtLoc.Text)) return;
                
                // 呼叫更新後的 SQLite 儲存方法
                DataManager.SaveInspectionRecord(
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), 
                    _txtLoc.Text, 
                    _txtUser.Text, 
                    _cmbStatus.Text
                );
                
                MessageBox.Show("資料已存入 SQLite 資料庫！", "系統提示");
                _txtLoc.Clear(); 
                _txtUser.Clear();
            };
            pnl.Controls.Add(btn);

            return pnl;
        }

        private void AddRow(Panel p, string labelText, Control inputControl, int y)
        {
            Label lbl = new Label { 
                Text = labelText, 
                Font = new Font("Microsoft JhengHei UI", 14), 
                Location = new Point(20, y), 
                AutoSize = true 
            };
            p.Controls.Add(lbl);
            inputControl.Location = new Point(180, y);
            inputControl.Width = 350;
            inputControl.Font = new Font("Microsoft JhengHei UI", 14);
            p.Controls.Add(inputControl);
        }
    }
}
