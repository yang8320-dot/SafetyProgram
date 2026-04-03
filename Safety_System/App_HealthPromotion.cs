using System;
using System.Windows.Forms;
using System.Drawing;

namespace Safety_System
{
    public class App_HealthPromotion
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill };
            Label info = new Label { Text = "💪 健康促進活動紀錄", Font = new Font("Microsoft JhengHei UI", 16F), AutoSize = true };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White };
            dgv.Columns.Add("Event", "活動/項目名稱");
            dgv.Columns.Add("Date", "執行日期");
            dgv.Columns.Add("Attendants", "參與人數");

            main.Controls.Add(info);
            main.Controls.Add(dgv);
            return main;
        }
    }
}
