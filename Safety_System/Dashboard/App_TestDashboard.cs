using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_TestDashboard
    {
        public Control GetView()
        {
            // 建立主面板
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            // 建立置中的文字標籤
            Label lblWelcome = new Label
            {
                Text = "檢測數據綜合管理看板\n\n\n請由上方【檢測數據】選單，切換至您欲操作的檢測項目。\n\n包含：環境監測、水質檢驗、TCLP、水錶校正等...",
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Regular),
                ForeColor = Color.DimGray,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };

            mainPanel.Controls.Add(lblWelcome);

            return mainPanel;
        }
    }
}
