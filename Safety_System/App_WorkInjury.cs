using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WorkInjury
    {
        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label lbl = new Label { Text = "🩹 工傷事件紀錄表", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), AutoSize = true, Margin = new Padding(10) };
            
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
            dgv.Columns.Add("Date", "受傷日期");
            dgv.Columns.Add("EmpName", "姓名");
            dgv.Columns.Add("Part", "受傷部位");
            dgv.Columns.Add("Severity", "嚴重程度");
            
            main.Controls.Add(lbl, 0, 0);
            main.Controls.Add(dgv, 0, 1);
            return main;
        }
    }
}
