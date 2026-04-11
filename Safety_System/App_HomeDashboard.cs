/// FILE: Safety_System/App_HomeDashboard.cs ///
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_HomeDashboard
    {
        public Control GetView()
        {
            Panel main = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, AutoScroll = true };

            // 標題區
            Label lblTitle = new Label {
                Text = "Safety System 綜合數據看板系統",
                Font = new Font("Microsoft JhengHei UI", 26F, FontStyle.Bold),
                ForeColor = Color.FromArgb(45, 62, 80),
                Location = new Point(40, 40),
                AutoSize = true
            };
            
            Label lblSubTitle = new Label {
                Text = "歡迎使用！請由上方選單進行詳細操作，或點選下方快捷按鈕進入常用模組。",
                Font = new Font("Microsoft JhengHei UI", 14F),
                ForeColor = Color.Gray,
                Location = new Point(45, 100),
                AutoSize = true
            };

            // 快捷按鈕容器
            FlowLayoutPanel flp = new FlowLayoutPanel {
                Location = new Point(40, 160),
                Size = new Size(1250, 600),
                BackColor = Color.Transparent
            };

            // 建立導覽大按鈕
            flp.Controls.Add(CreateShortcut("🛡️ 工安看板", "查看零災害天數與工安指標", Color.SteelBlue, () => new App_SafetyDashboard().GetView()));
            flp.Controls.Add(CreateShortcut("💧 水資源分析", "水情摘要、用水量與YoY比較圖表", Color.Teal, () => new App_WaterDashboard().GetView()));
            
            // 🟢 修正點：將 App_WaterTreatment 改為呼叫專屬水資源共用模組 App_Water_Generic
            flp.Controls.Add(CreateShortcut("📋 廢水水量紀錄", "每日廢水處理水量與讀數填報", Color.SeaGreen, () => new App_Water_Generic("Water", "WaterMeterReadings", "【日】廢水處理水量記錄").GetView()));
            
            flp.Controls.Add(CreateShortcut("🧪 檢測數據", "環境、飲用水、廢水檢測數據管理", Color.Chocolate, () => new App_TestDashboard().GetView()));
            flp.Controls.Add(CreateShortcut("♻️ 廢棄物月報", "紀錄廢棄物產出代碼、名稱與重量", Color.DimGray, () => new App_GenericTable("Waste", "WasteMonthly", "廢棄物統計表").GetView()));
            flp.Controls.Add(CreateShortcut("⚙️ 操作說明", "系統操作導覽、快捷鍵與密碼提示", Color.MediumPurple, () => new App_Instruction().GetView()));

            main.Controls.Add(lblTitle);
            main.Controls.Add(lblSubTitle);
            main.Controls.Add(flp);

            return main;
        }

        private Control CreateShortcut(string title, string desc, Color themeColor, Func<Control> loadFunc)
        {
            Panel p = new Panel { Size = new Size(380, 180), Margin = new Padding(0, 0, 30, 30), BackColor = Color.WhiteSmoke, Cursor = Cursors.Hand };
            
            // 繪製細邊框
            p.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, Color.Gainsboro, ButtonBorderStyle.Solid);
            
            Label l1 = new Label { Text = title, Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = themeColor, Location = new Point(20, 25), AutoSize = true };
            Label l2 = new Label
