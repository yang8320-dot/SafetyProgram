using System;
using System.Windows.Forms;
using System.Drawing;

namespace Safety_System
{
    public class App_WorkInjuryReport
    {
        public Control GetView()
        {
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White };
            dgv.Columns.Add("Date", "事故日期");
            dgv.Columns.Add("EmpId", "工號");
            dgv.Columns.Add("Description", "事故簡述");
            dgv.Columns.Add("DaysLost", "損失日數");
            return dgv;
        }
    }
}
