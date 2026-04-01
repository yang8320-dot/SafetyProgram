using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    /// <summary>
    /// 工安 - 巡檢模組：已針對 1440x810 解析度優化排版
    /// </summary>
    public class App_SafetyInspection
    {
        private TextBox _txtLocation = new TextBox();
        private TextBox _txtInspector = new TextBox();
        private ComboBox _cmbStatus = new ComboBox();

        /// <summary>
        /// 獲取優化後的巡檢介面
        /// </summary>
        public Control GetView()
        {
            Panel panel = new Panel { Dock = DockStyle.Fill };

            // 模組標題 (加大字體)
            Label lblTitle = new Label
            {
                Text = "[ 工安巡檢紀錄作業 ]",
                Font = new Font("Microsoft JhengHei", 20, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(20, 20)
            };
            panel.Controls.Add(lblTitle);

            // 使用加大間距與字體的排版
            AddRow(panel, "巡檢地點:", _txtLocation, 100);
            AddRow(panel, "巡檢人員:", _txtInspector, 160);
            
            // 設備狀態下拉選單
            Label lblStatus = new Label
            {
                Text = "設備狀態:",
                Font = new Font("Microsoft JhengHei", 14),
                Location = new Point(20, 220),
                AutoSize = true
            };
            _cmbStatus.Location = new Point(180, 220);
            _cmbStatus.Width = 350;
            _cmbStatus.Font = new Font("Microsoft JhengHei", 14);
            _cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbStatus.Items.AddRange(new string[] { "正常 (Normal)", "異常 (Abnormal)", "待派修 (Repair Needed)" });
            _cmbStatus.SelectedIndex = 0;
            
            panel.Controls.Add(lblStatus);
            panel.Controls.Add(_cmbStatus);

            // 儲存按鈕 (加大尺寸)
            Button btnSave = new Button
            {
                Text = "儲存紀錄",
                Location = new Point(180, 300),
                Size = new Size(180, 60),
                BackColor = Color.LightSteelBlue,
                Font = new Font("Microsoft JhengHei", 14, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.Click += BtnSave_Click;
            panel.Controls.Add(btnSave);

            return panel;
        }

        /// <summary>
        /// 輔助方法：快速建立標籤與輸入框組合
        /// </summary>
        private void AddRow(Panel p, string labelText, Control inputControl, int y)
        {
            Label lbl = new Label
            {
                Text = labelText,
                Font = new Font("Microsoft JhengHei", 14),
                Location = new Point(20, y),
                AutoSize = true
            };
            p.Controls.Add(lbl);

            inputControl.Location = new Point(180, y);
            inputControl.Width = 350;
            inputControl.Font = new Font("Microsoft JhengHei(微軟正黑體)", 14);
            p.Controls.Add(inputControl);
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtLocation.Text) || string.IsNullOrWhiteSpace(_txtInspector.Text))
            {
                MessageBox.Show("請完整填寫巡檢地點與人員名稱。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string status = _cmbStatus.SelectedItem?.ToString() ?? "未設定";

            // 儲存至純文字檔
            DataManager.SaveInspectionRecord(date, _txtLocation.Text, _txtInspector.Text, status);

            MessageBox.Show("巡檢紀錄已成功寫入資料檔。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            // 清空欄位
            _txtLocation.Clear();
            _txtInspector.Clear();
            _cmbStatus.SelectedIndex = 0;
        }
    }
}
