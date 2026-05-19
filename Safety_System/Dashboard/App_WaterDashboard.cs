/// FILE: Safety_System/App_WaterDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterDashboard
    {
        // 每日查詢區間
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        
        // 每月查詢區間
        private ComboBox _cboStartMonthYear, _cboStartMonthMonth;
        private ComboBox _cboEndMonthYear, _cboEndMonthMonth;

        // 小標題參考
        private Label _lblBox2Sub1, _lblBox2Sub2, _lblBox2Sub3, _lblBox2Sub4; 
        private Label _lblBox3Sub1, _lblBox3Sub2, _lblBox3Sub3, _lblBox3Sub4; 
        private Label _lblBox4Sub1, _lblBox4Sub2, _lblBox4Sub3, _lblBox4Sub4; 
        private Label _lblBox5Sub1, _lblBox5Sub2, _lblBox5Sub3, _lblBox5Sub4; 
        private Label _lblBox6Sub1, _lblBox6Sub2, _lblBox6Sub3, _lblBox6Sub4; 
        private Label _lblBox7Sub1, _lblBox7Sub2, _lblBox7Sub3, _lblBox7Sub4; 

        // 數據欄位參考
        private Panel _pnlBox2Data1, _pnlBox2Data2, _pnlBox2Data3, _pnlBox2Data4;
        private Panel _pnlBox3Data1, _pnlBox3Data2, _pnlBox3Data3, _pnlBox3Data4;
        private Panel _pnlBox4Data1, _pnlBox4Data2, _pnlBox4Data3, _pnlBox4Data4;
        private Panel _pnlBox5Data1, _pnlBox5Data2, _pnlBox5Data3, _pnlBox5Data4;
        private Panel _pnlBox6Data1, _pnlBox6Data2, _pnlBox6Data3, _pnlBox6Data4;
        private Panel _pnlBox7Data1, _pnlBox7Data2, _pnlBox7Data3, _pnlBox7Data4;
        
        // 模組外框參考 (供 PDF 使用)
        private Panel _pnlWaterBox, _pnlRecycleBox, _pnlChemBox, _pnlDailyUsageBox;
        private Panel _pnlDischargeBox, _pnlMonthlyVolumeBox;
        
        private Panel _mainScrollPanel;
        private const string DbName = "Water";
        private readonly string SettingsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WaterSettings.txt");

        // 儲存各模組的顯示設定
        private Dictionary<string, bool> _visibilitySettings = new Dictionary<string, bool>();

        // 🟢 供自定義單位的全域快取
        public static Dictionary<string, string> CustomUnitsCache = new Dictionary<string, string>();

        public Control GetView()
        {
            CustomStatEngine.LoadSettings(); 
            LoadVisibilitySettings(); 

            _mainScrollPanel = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.WhiteSmoke, 
                AutoScroll = true, 
                Padding = new Padding(20) 
            };

            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Top, 
                AutoSize = true, 
                ColumnCount = 1, 
                RowCount = 8 
            };

            // ==========================================
            // 大框 1：日報表功能選單與日期查詢
            // ==========================================
            Panel box1 = new Panel { Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            box1.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, box1.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            FlowLayoutPanel flpTop = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, Padding = new Padding(15) };
            Label lblTitle = new Label { Text = "💧 水資源綜合數據看板 (日報表區)", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            
            _cboStartYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboStartDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndDay = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitDateComboBoxes(); 

            Button btnSearchDaily = new Button { Text = "🔍 查詢日統計", Size = new Size(140, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSearchDaily.Click += (s, e) => LoadDailyData();
            
            Button btnPdf = new Button { Text = "📄 選擇並導出 PDF", Size = new Size(180, 32), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnPdf.Click += BtnPdf_Click;

            flpControls.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboStartYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboEndYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndDay, new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearchDaily, 
                btnPdf
            });

            flpTop.Controls.Add(lblTitle); 
            flpTop.Controls.Add(flpControls); 
            box1.Controls.Add(flpTop);
            mainLayout.Controls.Add(box1, 0, 0);

            // 日報表資料框
            _pnlWaterBox = BuildNineGridBox("WaterStat", "台灣玻璃彰濱廠 - 水資源數據統計", Color.Teal, out _lblBox2Sub1, out _lblBox2Sub2, out _lblBox2Sub3, out _lblBox2Sub4, out _pnlBox2Data1, out _pnlBox2Data2, out _pnlBox2Data3, out _pnlBox2Data4);
            _pnlRecycleBox = BuildNineGridBox("RecycleStat", "台灣玻璃彰濱廠 - 回收水統計", Color.ForestGreen, out _lblBox3Sub1, out _lblBox3Sub2, out _lblBox3Sub3, out _lblBox3Sub4, out _pnlBox3Data1, out _pnlBox3Data2, out _pnlBox3Data3, out _pnlBox3Data4);
            _pnlChemBox = BuildNineGridBox("ChemStat", "台灣玻璃彰濱廠 - 藥劑數據統計", Color.Sienna, out _lblBox4Sub1, out _lblBox4Sub2, out _lblBox4Sub3, out _lblBox4Sub4, out _pnlBox4Data1, out _pnlBox4Data2, out _pnlBox4Data3, out _pnlBox4Data4);
            _pnlDailyUsageBox = BuildNineGridBox("DailyUsage", "台灣玻璃彰濱廠 - 自來水用量統計", Color.MediumBlue, out _lblBox5Sub1, out _lblBox5Sub2, out _lblBox5Sub3, out _lblBox5Sub4, out _pnlBox5Data1, out _pnlBox5Data2, out _pnlBox5Data3, out _pnlBox5Data4);
            
            mainLayout.Controls.Add(_pnlWaterBox, 0, 1);
            mainLayout.Controls.Add(_pnlRecycleBox, 0, 2);
            mainLayout.Controls.Add(_pnlChemBox, 0, 3);
            mainLayout.Controls.Add(_pnlDailyUsageBox, 0, 4); 

            // ==========================================
            // 月報表功能選單與日期查詢
            // ==========================================
            Panel boxMonthlyFilter = new Panel { Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 100), BackColor = Color.White, Margin = new Padding(0, 20, 0, 20) };
            boxMonthlyFilter.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, boxMonthlyFilter.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            FlowLayoutPanel flpTopMonthly = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, Padding = new Padding(15) };
            Label lblTitleMonthly = new Label { Text = "💧 月度結算數據看板 (月報表區)", Font = new Font("Microsoft JhengHei UI", 24F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            FlowLayoutPanel flpControlsMonthly = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            
            _cboStartMonthYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboStartMonthMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };
            _cboEndMonthYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 80 };
            _cboEndMonthMonth = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Width = 60 };

            InitMonthlyComboBoxes();

            Button btnSearchMonthly = new Button { Text = "🔍 查詢月統計", Size = new Size(140, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(15, 0, 0, 0) };
            btnSearchMonthly.Click += (s, e) => LoadMonthlyData();

            flpControlsMonthly.Controls.AddRange(new Control[] { 
                new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboStartMonthYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboStartMonthMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                new Label { Text = "~", AutoSize = true, Margin = new Padding(0, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) },
                _cboEndMonthYear, new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                _cboEndMonthMonth, new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) },
                btnSearchMonthly
            });

            flpTopMonthly.Controls.Add(lblTitleMonthly); 
            flpTopMonthly.Controls.Add(flpControlsMonthly); 
            boxMonthlyFilter.Controls.Add(flpTopMonthly);
            mainLayout.Controls.Add(boxMonthlyFilter, 0, 5);

            // 月報表資料框
            _pnlDischargeBox = BuildNineGridBox("DischargeStat", "台灣玻璃彰濱廠 - 納管排放數據統計", Color.DarkCyan, out _lblBox6Sub1, out _lblBox6Sub2, out _lblBox6Sub3, out _lblBox6Sub4, out _pnlBox6Data1, out _pnlBox6Data2, out _pnlBox6Data3, out _pnlBox6Data4);
            _pnlMonthlyVolumeBox = BuildNineGridBox("MonthlyVolume", "台灣玻璃彰濱廠 - 自來水用量(繳費單)月統計", Color.MediumPurple, out _lblBox7Sub1, out _lblBox7Sub2, out _lblBox7Sub3, out _lblBox7Sub4, out _pnlBox7Data1, out _pnlBox7Data2, out _pnlBox7Data3, out _pnlBox7Data4);

            mainLayout.Controls.Add(_pnlDischargeBox, 0, 6);
            mainLayout.Controls.Add(_pnlMonthlyVolumeBox, 0, 7);

            _mainScrollPanel.Controls.Add(mainLayout);
            LoadDailyData();   
            LoadMonthlyData(); 
            return _mainScrollPanel;
        }

        private void LoadVisibilitySettings()
        {
            _visibilitySettings.Clear();
            if (File.Exists(SettingsFile)) 
            {
                try 
                {
                    foreach (var line in File.ReadAllLines(SettingsFile, Encoding.UTF8)) 
                    {
                        var parts = line.Split('|');
                        if (parts.Length == 2) 
                        {
                            _visibilitySettings[parts[0]] = parts[1] == "1";
                        }
                    }
                } 
                catch { }
            }
        }

        private void SaveVisibilitySettings()
        {
            try 
            {
                var lines = _visibilitySettings.Select(kvp => $"{kvp.Key}|{(kvp.Value ? "1" : "0")}").ToArray();
                File.WriteAllLines(SettingsFile, lines, Encoding.UTF8);
            } 
            catch { }
        }

        private bool IsVisible(string module, string key)
        {
            string dictKey = $"{module}_{key}";
            if (_visibilitySettings.ContainsKey(dictKey)) 
            {
                return _visibilitySettings[dictKey];
            }

            // 預設隱藏邏輯
            if (key.Contains("貯存區") || key.Contains("清水池") || key.Contains("BOD") || key.Contains("氨氮")) 
            {
                _visibilitySettings[dictKey] = false;
                return false;
            }

            _visibilitySettings[dictKey] = true;
            return true;
        }

        private void OpenSettingsDialog(string moduleName, Dictionary<string, double> currentData)
        {
            using (Form f = new Form() { Width = 350, Height = 450, Text = "顯示欄位設定", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選此區塊要顯示的數據項目：", Dock = DockStyle.Top, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(10), BorderStyle = BorderStyle.None, BackColor = f.BackColor };
                
                var keys = currentData.Keys.Where(k => !k.Contains("【") && !k.EndsWith("/M3") && k != "PAC" && k != "NAOH" && k != "高分子").ToList();
                
                if (moduleName == "ChemStat") 
                {
                    keys = new List<string> { "PAC", "NAOH", "高分子", "PAC/M3", "NAOH/M3", "高分子/M3" };
                } 
                else if (moduleName == "DischargeStat") 
                {
                    keys = currentData.Keys.ToList();
                }

                foreach (var k in keys) 
                {
                    string dictKey = $"{moduleName}_{k}";
                    bool isChecked = IsVisible(moduleName, k);
                    clb.Items.Add(k, isChecked);
                }
                
                f.Controls.Add(clb);

                Button btnOk = new Button { Text = "💾 儲存並套用", Dock = DockStyle.Bottom, Height = 45, DialogResult = DialogResult.OK, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                f.Controls.Add(btnOk);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    for (int i = 0; i < clb.Items.Count; i++) 
                    {
                        string k = clb.Items[i].ToString();
                        _visibilitySettings[$"{moduleName}_{k}"] = clb.GetItemChecked(i);
                    }
                    SaveVisibilitySettings();
                    
                    if (moduleName == "DischargeStat" || moduleName == "MonthlyVolume") LoadMonthlyData();
                    else LoadDailyData();
                }
            }
        }

        private void InitDateComboBoxes()
        {
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) 
            { 
                _cboStartYear.Items.Add(i); 
                _cboEndYear.Items.Add(i); 
            }
            
            for (int i = 1; i <= 12; i++) 
            { 
                _cboStartMonth.Items.Add(i.ToString("D2")); 
                _cboEndMonth.Items.Add(i.ToString("D2")); 
            }
            
            _cboStartYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboStartMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            _cboEndYear.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            _cboEndMonth.SelectedIndexChanged += (s, e) => UpdateDaysCombo(_cboEndYear, _cboEndMonth, _cboEndDay);
            
            SetComboValue(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-1));
            SetComboValue(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);
        }

        private void InitMonthlyComboBoxes()
        {
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) 
            { 
                _cboStartMonthYear.Items.Add(i); 
                _cboEndMonthYear.Items.Add(i); 
            }
            
            for (int i = 1; i <= 12; i++) 
            { 
                _cboStartMonthMonth.Items.Add(i.ToString("D2")); 
                _cboEndMonthMonth.Items.Add(i.ToString("D2")); 
            }
            
            DateTime start = DateTime.Today.AddMonths(-6);
            _cboStartMonthYear.SelectedItem = start.Year; 
            _cboStartMonthMonth.SelectedItem = start.Month.ToString("D2");
            _cboEndMonthYear.SelectedItem = DateTime.Today.Year; 
            _cboEndMonthMonth.SelectedItem = DateTime.Today.Month.ToString("D2");
        }

        private void UpdateDaysCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null) return;
            
            int days = DateTime.DaysInMonth((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()));
            string currentDay = d.SelectedItem?.ToString();
            
            d.Items.Clear();
            for (int i = 1; i <= days; i++) d.Items.Add(i.ToString("D2"));
            
            if (currentDay != null && d.Items.Contains(currentDay)) 
            {
                d.SelectedItem = currentDay;
            }
            else 
            {
                d.SelectedIndex = d.Items.Count - 1; 
            }
        }

        private void SetComboValue(ComboBox y, ComboBox m, ComboBox d, DateTime date)
        {
            y.SelectedItem = date.Year; 
            m.SelectedItem = date.Month.ToString("D2");
            UpdateDaysCombo(y, m, d); 
            d.SelectedItem = date.Day.ToString("D2");
        }

        private DateTime GetDateFromCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            int day = int.Parse(d.SelectedItem.ToString());
            int maxDay = DateTime.DaysInMonth((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()));
            return new DateTime((int)y.SelectedItem, int.Parse(m.SelectedItem.ToString()), day > maxDay ? maxDay : day);
        }

        private Panel BuildNineGridBox(string moduleName, string mainTitle, Color headerColor, out Label l1, out Label l2, out Label l3, out Label l4, out Panel d1, out Panel d2, out Panel d3, out Panel d4)
        {
            Panel outer = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.White, Margin = new Padding(0, 0, 0, 20) };
            outer.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, outer.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            TableLayoutPanel grid = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, RowCount = 3, Padding = new Padding(10) };
            for (int i = 0; i < 4; i++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F)); 
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));      

            Panel pnlHeader = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            Label lblMainTitle = new Label { Text = mainTitle, Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };
            
            Button btnSettings = new Button { Text = "⚙️ 設定顯示", Size = new Size(110, 30), BackColor = Color.LightGray, ForeColor = Color.Black, Font = new Font("Microsoft JhengHei UI", 10F), Dock = DockStyle.Right, Cursor = Cursors.Hand };
            Button btnCustom = new Button { Text = "➕ 統計設定", Size = new Size(110, 30), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 10F), Dock = DockStyle.Right, Cursor = Cursors.Hand, Margin = new Padding(0,0,10,0) };

            btnSettings.Tag = moduleName;
            btnSettings.Click += (s, e) => {
                string mod = ((Button)s).Tag.ToString();
                Dictionary<string, double> currentData = new Dictionary<string, double>();
                
                if (mod == "DailyUsage") currentData = CalculateDailyUsageStats(GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay).ToString("yyyy-MM-dd"), GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay).ToString("yyyy-MM-dd"));
                else if (mod == "MonthlyVolume") currentData = CalculateMonthlyVolumeStats($"{_cboStartMonthYear.Text}-{_cboStartMonthMonth.Text}", $"{_cboEndMonthYear.Text}-{_cboEndMonthMonth.Text}");
                else if (mod == "ChemStat") currentData = CalculateChemicalStats(GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay).ToString("yyyy-MM-dd"), GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay).ToString("yyyy-MM-dd"));
                else if (mod == "RecycleStat") currentData = CalculateRecycleStats(GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay).ToString("yyyy-MM-dd"), GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay).ToString("yyyy-MM-dd"));
                else if (mod == "DischargeStat") currentData = CalculateDischargeStats($"{_cboStartMonthYear.Text}-{_cboStartMonthMonth.Text}", $"{_cboEndMonthYear.Text}-{_cboEndMonthMonth.Text}");
                else currentData = ProcessBox2Data(GetSumsEndingWith(GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay).ToString("yyyy-MM-dd"), GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay).ToString("yyyy-MM-dd"), "WaterMeterReadings", "WaterUsageDaily"));
                
                var customs = CustomStatEngine.GetStatsForModule(mod);
                foreach (var c in customs) currentData[c.StatName] = 0;

                OpenSettingsDialog(mod, currentData);
            };

            btnCustom.Tag = moduleName;
            btnCustom.Click += (s, e) => {
                CustomStatEngine.OpenConfigDialog(moduleName);
                if (moduleName == "DischargeStat" || moduleName == "MonthlyVolume") LoadMonthlyData();
                else LoadDailyData();
            };

            pnlHeader.Controls.Add(btnSettings);
            pnlHeader.Controls.Add(btnCustom);
            pnlHeader.Controls.Add(lblMainTitle);

            grid.Controls.Add(pnlHeader, 0, 0); 
            grid.SetColumnSpan(pnlHeader, 4);

            l1 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            l2 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            l3 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };
            l4 = new Label { Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.White, BackColor = headerColor, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(2) };

            grid.Controls.Add(l1, 0, 1); 
            grid.Controls.Add(l2, 1, 1); 
            grid.Controls.Add(l3, 2, 1); 
            grid.Controls.Add(l4, 3, 1);

            d1 = CreateDataPanel(); 
            d2 = CreateDataPanel(); 
            d3 = CreateDataPanel(); 
            d4 = CreateDataPanel();
            
            grid.Controls.Add(d1, 0, 2); 
            grid.Controls.Add(d2, 1, 2); 
            grid.Controls.Add(d3, 2, 2); 
            grid.Controls.Add(d4, 3, 2);

            outer.Controls.Add(grid);
            return outer;
        }

        private Panel CreateDataPanel()
        {
            return new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, MinimumSize = new Size(0, 100), FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.FromArgb(248, 249, 250), Margin = new Padding(2), Padding = new Padding(10) };
        }

        private void LoadDailyData()
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
            DateTime dtS = GetDateFromCombo(_cboStartYear, _cboStartMonth, _cboStartDay);
            DateTime dtE = GetDateFromCombo(_cboEndYear, _cboEndMonth, _cboEndDay);

            UpdateSubtitles(_lblBox2Sub1, _lblBox2Sub2, _lblBox2Sub3, _lblBox2Sub4, dtS, dtE, false);
            UpdateSubtitles(_lblBox3Sub1, _lblBox3Sub2, _lblBox3Sub3, _lblBox3Sub4, dtS, dtE, true);
            UpdateSubtitles(_lblBox4Sub1, _lblBox4Sub2, _lblBox4Sub3, _lblBox4Sub4, dtS, dtE, false);
            UpdateSubtitles(_lblBox5Sub1, _lblBox5Sub2, _lblBox5Sub3, _lblBox5Sub4, dtS, dtE, false); 

            string sS = dtS.ToString("yyyy-MM-dd"), sE = dtE.ToString("yyyy-MM-dd");
            string sLY_S = dtS.AddYears(-1).ToString("yyyy-MM-dd"), sLY_E = dtE.AddYears(-1).ToString("yyyy-MM-dd");
            string sL2Y_S = dtS.AddYears(-2).ToString("yyyy-MM-dd"), sL2Y_E = dtE.AddYears(-2).ToString("yyyy-MM-dd");

            FillDataPanels("WaterStat", _pnlBox2Data1, _pnlBox2Data2, _pnlBox2Data3, _pnlBox2Data4, 
                MergeCustom("WaterStat", sS, sE, ProcessBox2Data(GetSumsEndingWith(sS, sE, "WaterMeterReadings", "WaterUsageDaily"))), 
                MergeCustom("WaterStat", sLY_S, sLY_E, ProcessBox2Data(GetSumsEndingWith(sLY_S, sLY_E, "WaterMeterReadings", "WaterUsageDaily"))), 
                MergeCustom("WaterStat", sL2Y_S, sL2Y_E, ProcessBox2Data(GetSumsEndingWith(sL2Y_S, sL2Y_E, "WaterMeterReadings", "WaterUsageDaily"))));

            FillDataPanels("RecycleStat", _pnlBox3Data1, _pnlBox3Data2, _pnlBox3Data3, _pnlBox3Data4, 
                MergeCustom("RecycleStat", sS, sE, CalculateRecycleStats(sS, sE)), 
                MergeCustom("RecycleStat", sLY_S, sLY_E, CalculateRecycleStats(sLY_S, sLY_E)), 
                MergeCustom("RecycleStat", sL2Y_S, sL2Y_E, CalculateRecycleStats(sL2Y_S, sL2Y_E)), true);

            FillDataPanels("ChemStat", _pnlBox4Data1, _pnlBox4Data2, _pnlBox4Data3, _pnlBox4Data4, 
                MergeCustom("ChemStat", sS, sE, CalculateChemicalStats(sS, sE)), 
                MergeCustom("ChemStat", sLY_S, sLY_E, CalculateChemicalStats(sLY_S, sLY_E)), 
                MergeCustom("ChemStat", sL2Y_S, sL2Y_E, CalculateChemicalStats(sL2Y_S, sL2Y_E)));

            FillDataPanels("DailyUsage", _pnlBox5Data1, _pnlBox5Data2, _pnlBox5Data3, _pnlBox5Data4, 
                MergeCustom("DailyUsage", sS, sE, CalculateDailyUsageStats(sS, sE)), 
                MergeCustom("DailyUsage", sLY_S, sLY_E, CalculateDailyUsageStats(sLY_S, sLY_E)), 
                MergeCustom("DailyUsage", sL2Y_S, sL2Y_E, CalculateDailyUsageStats(sL2Y_S, sL2Y_E)));

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private void LoadMonthlyData()
        {
            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
            string startYM = $"{_cboStartMonthYear.Text}-{_cboStartMonthMonth.Text}";
            string endYM = $"{_cboEndMonthYear.Text}-{_cboEndMonthMonth.Text}";
            
            string startLY = $"{int.Parse(_cboStartMonthYear.Text) - 1}-{_cboStartMonthMonth.Text}";
            string endLY = $"{int.Parse(_cboEndMonthYear.Text) - 1}-{_cboEndMonthMonth.Text}";
            string startL2Y = $"{int.Parse(_cboStartMonthYear.Text) - 2}-{_cboStartMonthMonth.Text}";
            string endL2Y = $"{int.Parse(_cboEndMonthYear.Text) - 2}-{_cboEndMonthMonth.Text}";

            _lblBox6Sub1.Text = $"【{startYM.Replace("-","/")} ~ {endYM.Replace("-","/")}】\n區間數據統計";
            _lblBox6Sub2.Text = $"【{startLY.Replace("-","/")} ~ {endLY.Replace("-","/")}】\n去年同期區間數據統計";
            _lblBox6Sub3.Text = $"【{startL2Y.Replace("-","/")} ~ {endL2Y.Replace("-","/")}】\n前年同期區間數據統計";
            _lblBox6Sub4.Text = $"【{startYM.Replace("-","/")} ~ {endYM.Replace("-","/")}】\n與去年同期差異分析";

            _lblBox7Sub1.Text = $"【{startYM.Replace("-","/")} ~ {endYM.Replace("-","/")}】\n區間數據統計";
            _lblBox7Sub2.Text = $"【{startLY.Replace("-","/")} ~ {endLY.Replace("-","/")}】\n去年同期區間數據統計";
            _lblBox7Sub3.Text = $"【{startL2Y.Replace("-","/")} ~ {endL2Y.Replace("-","/")}】\n前年同期區間數據統計";
            _lblBox7Sub4.Text = $"【{startYM.Replace("-","/")} ~ {endYM.Replace("-","/")}】\n與去年同期差異分析";

            // 🟢 注意：這裡傳入的格式已經是 yyyy-MM，可以對應到資料庫內的年月欄位了
            FillDataPanels("DischargeStat", _pnlBox6Data1, _pnlBox6Data2, _pnlBox6Data3, _pnlBox6Data4, 
                MergeCustom("DischargeStat", startYM, endYM, CalculateDischargeStats(startYM, endYM)), 
                MergeCustom("DischargeStat", startLY, endLY, CalculateDischargeStats(startLY, endLY)), 
                MergeCustom("DischargeStat", startL2Y, endL2Y, CalculateDischargeStats(startL2Y, endL2Y)));

            FillDataPanels("MonthlyVolume", _pnlBox7Data1, _pnlBox7Data2, _pnlBox7Data3, _pnlBox7Data4, 
                MergeCustom("MonthlyVolume", startYM, endYM, CalculateMonthlyVolumeStats(startYM, endYM)), 
                MergeCustom("MonthlyVolume", startLY, endLY, CalculateMonthlyVolumeStats(startLY, endLY)), 
                MergeCustom("MonthlyVolume", startL2Y, endL2Y, CalculateMonthlyVolumeStats(startL2Y, endL2Y)));

            if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
        }

        private Dictionary<string, double> MergeCustom(string module, string start, string end, Dictionary<string, double> current)
        {
            var customs = CustomStatEngine.EvaluateForModule(module, start, end);
            foreach (var kvp in customs) current[kvp.Key] = kvp.Value;
            return current;
        }

        private void UpdateSubtitles(Label l1, Label l2, Label l3, Label l4, DateTime dtS, DateTime dtE, bool isRecycle)
        {
            string suffix = isRecycle ? "回收水量統計" : "數據統計";
            l1.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n區間{suffix}";
            l2.Text = $"【{dtS.AddYears(-1):yyyy/MM/dd} ~ {dtE.AddYears(-1):yyyy/MM/dd}】\n去年同期區間{suffix}";
            l3.Text = $"【{dtS.AddYears(-2):yyyy/MM/dd} ~ {dtE.AddYears(-2):yyyy/MM/dd}】\n前年同期區間{suffix}";
            l4.Text = $"【{dtS:yyyy/MM/dd} ~ {dtE:yyyy/MM/dd}】\n與去年同期差異分析"; 
        }

        private Dictionary<string, double> CalculateDailyUsageStats(string start, string end)
        {
            var dict = new Dictionary<string, double> {
                { "廠區", 0 }, { "行政區", 0 }, { "全廠用量", 0 },
                { "廠區平均", 0 }, { "行政區平均", 0 }, { "全廠平均", 0 },
                { "自來水至貯存區", 0 }, { "自來水至貯存區平均", 0 },
                { "自來水至清水池", 0 }, { "自來水至清水池平均", 0 } 
            };

            DataTable dt = null;
            try { dt = DataManager.GetTableData(DbName, "WaterUsageDaily", "日期", start, end); } catch { return dict; }
            if (dt == null) return dict;

            int validDays = 0;
            foreach (DataRow r in dt.Rows) 
            {
                bool hasData = false;
                if (dt.Columns.Contains("廠區自來水使用量") && double.TryParse(r["廠區自來水使用量"]?.ToString().Replace(",", ""), out double f)) { dict["廠區"] += f; hasData = true; }
                if (dt.Columns.Contains("行政區自來水使用量") && double.TryParse(r["行政區自來水使用量"]?.ToString().Replace(",", ""), out double a)) { dict["行政區"] += a; hasData = true; }
                
                if (dt.Columns.Contains("自來水至貯存池日統計") && double.TryParse(r["自來水至貯存池日統計"]?.ToString().Replace(",", ""), out double s)) { dict["自來水至貯存區"] += s; hasData = true; }
                if (dt.Columns.Contains("自來水至清水池日統計") && double.TryParse(r["自來水至清水池日統計"]?.ToString().Replace(",", ""), out double c)) { dict["自來水至清水池"] += c; hasData = true; }
                if (hasData) validDays++;
            }

            dict["全廠用量"] = dict["廠區"] + dict["行政區"];
            
            if (validDays > 0) 
            {
                dict["廠區平均"] = dict["廠區"] / validDays;
                dict["行政區平均"] = dict["行政區"] / validDays;
                dict["全廠平均"] = dict["全廠用量"] / validDays;
                dict["自來水至貯存區平均"] = dict["自來水至貯存區"] / validDays;
                dict["自來水至清水池平均"] = dict["自來水至清水池"] / validDays;
            }
            return dict;
        }

        // 🟢 改用 年月 查詢
        private Dictionary<string, double> CalculateMonthlyVolumeStats(string startYM, string endYM)
        {
            var dict = new Dictionary<string, double> { 
                { "廠區", 0 }, { "行政區", 0 }, { "全廠用量", 0 },
                { "廠區月平均", 0 }, { "行政區月平均", 0 }, { "全廠月平均", 0 } 
            };
            
            DataTable dt = null;
            try { dt = DataManager.GetTableData(DbName, "WaterVolume", "年月", startYM, endYM); } catch { return dict; }
            if (dt == null) return dict;

            int validMonths = 0; 
            foreach (DataRow r in dt.Rows) 
            {
                bool hasData = false;
                if (dt.Columns.Contains("廠區自來水繳費單") && double.TryParse(r["廠區自來水繳費單"]?.ToString().Replace(",", ""), out double f)) { dict["廠區"] += f; hasData = true; }
                if (dt.Columns.Contains("行政區自來水繳費單") && double.TryParse(r["行政區自來水繳費單"]?.ToString().Replace(",", ""), out double a)) { dict["行政區"] += a; hasData = true; }
                if (hasData) validMonths++;
            }

            dict["全廠用量"] = dict["廠區"] + dict["行政區"];

            if (validMonths > 0) 
            {
                dict["廠區月平均"] = dict["廠區"] / validMonths;
                dict["行政區月平均"] = dict["行政區"] / validMonths;
                dict["全廠月平均"] = dict["全廠用量"] / validMonths;
            }

            return dict;
        }

        // 🟢 改用 年月 查詢
        private Dictionary<string, double> CalculateDischargeStats(string startYM, string endYM)
        {
            var dict = new Dictionary<string, double>();
            dict["納管排放水量"] = 0;
            string[] metrics = { "SS", "COD", "BOD", "氨氮" };
            foreach (var m in metrics) 
            { 
                dict[$"{m}平均值"] = 0; 
                dict[$"{m}最大值"] = 0; 
                dict[$"{m}最小值"] = 0; 
            }

            DataTable dt = null;
            try { dt = DataManager.GetTableData(DbName, "DischargeData", "年月", startYM, endYM); } catch { return dict; }
            if (dt == null) return dict;

            double totalWater = 0;
            var lists = new Dictionary<string, List<double>>();
            foreach (var m in metrics) lists[m] = new List<double>();

            foreach (DataRow r in dt.Rows) 
            {
                if (dt.Columns.Contains("水量") && double.TryParse(r["水量"]?.ToString().Replace(",", ""), out double w)) totalWater += w;
                foreach (var m in metrics) 
                {
                    if (dt.Columns.Contains(m) && double.TryParse(r[m]?.ToString().Replace(",", ""), out double val)) 
                    {
                        lists[m].Add(val);
                    }
                }
            }

            dict["納管排放水量"] = totalWater;
            foreach (var m in metrics) 
            {
                if (lists[m].Count > 0) 
                {
                    dict[$"{m}平均值"] = lists[m].Average();
                    dict[$"{m}最大值"] = lists[m].Max();
                    dict[$"{m}最小值"] = lists[m].Min();
                }
            }
            return dict;
        }

        private Dictionary<string, double> GetSumsEndingWith(string start, string end, params string[] tableNames) 
        {
            var res = new Dictionary<string, double>();
            foreach (string tbl in tableNames) 
            {
                try 
                {
                    var dt = DataManager.GetTableData(DbName, tbl, "日期", start, end);
                    if (dt == null) continue;
                    
                    var targetCols = dt.Columns.Cast<DataColumn>().Where(c => c.ColumnName.EndsWith("日統計") || c.ColumnName == "污泥產出KG").Select(c => c.ColumnName).ToList();
                    foreach (DataRow r in dt.Rows) 
                    {
                        foreach (string col in targetCols) 
                        {
                            string cleanName = col == "污泥產出KG" ? "污泥量" : col.Replace("日統計", ""); 
                            if (!res.ContainsKey(cleanName)) res[cleanName] = 0;
                            if (double.TryParse(r[col]?.ToString().Replace(",", ""), out double v)) res[cleanName] += v;
                        }
                    }
                } 
                catch { }
            }
            return res;
        }

        private Dictionary<string, double> ProcessBox2Data(Dictionary<string, double> raw) 
        {
            var res = new Dictionary<string, double>();
            foreach (var kvp in raw) 
            {
                if (kvp.Key.Contains("回收水6吋") || kvp.Key.Contains("軟水") || kvp.Key.Contains("回收水雙介質")) continue;
                
                double val = kvp.Value;
                if (kvp.Key == "用電量") val = kvp.Value * 100;
                else if (kvp.Key == "污泥量") val = kvp.Value / 1000.0;
                
                res[kvp.Key] = val;
                if (kvp.Key == "濃縮水至逆洗池" || (kvp.Key == "濃縮水至冷卻水池" && !raw.ContainsKey("濃縮水至逆洗池"))) 
                {
                    res["濃縮水合計"] = (raw.ContainsKey("濃縮水至冷卻水池") ? raw["濃縮水至冷卻水池"] : 0) + (raw.ContainsKey("濃縮水至逆洗池") ? raw["濃縮水至逆洗池"] : 0);
                }
            }
            
            if (raw.ContainsKey("回收水雙介質A") || raw.ContainsKey("回收水雙介質B")) 
            {
                res["總回收水量"] = (raw.ContainsKey("回收水雙介質A") ? raw["回收水雙介質A"] : 0) + (raw.ContainsKey("回收水雙介質B") ? raw["回收水雙介質B"] : 0);
            }
            return res;
        }

        private Dictionary<string, double> CalculateRecycleStats(string start, string end) 
        {
            var dict = new Dictionary<string, double> { { "廢水處理量", 0 }, { "總回收水量", 0 }, { "回收率(%)", 0 } };
            double sumA = 0, sumB = 0;
            
            try 
            {
                var dt = DataManager.GetTableData(DbName, "WaterMeterReadings", "日期", start, end);
                if (dt != null) 
                {
                    foreach (DataRow r in dt.Rows) 
                    {
                        if (dt.Columns.Contains("廢水處理量日統計") && double.TryParse(r["廢水處理量日統計"]?.ToString().Replace(",", ""), out double w)) dict["廢水處理量"] += w;
                        if (dt.Columns.Contains("回收水雙介質A日統計") && double.TryParse(r["回收水雙介質A日統計"]?.ToString().Replace(",", ""), out double a)) sumA += a;
                        if (dt.Columns.Contains("回收水雙介質B日統計") && double.TryParse(r["回收水雙介質B日統計"]?.ToString().Replace(",", ""), out double b)) sumB += b;
                    }
                }
            } 
            catch { }
            
            dict["總回收水量"] = sumA + sumB;
            if (dict["廢水處理量"] > 0) dict["回收率(%)"] = (dict["總回收水量"] / dict["廢水處理量"]) * 100;
            return dict;
        }

        private Dictionary<string, double> CalculateChemicalStats(string start, string end) 
        {
            var dict = new Dictionary<string, double> { { "PAC", 0 }, { "NAOH", 0 }, { "高分子", 0 }, { "PAC/M3", 0 }, { "NAOH/M3", 0 }, { "高分子/M3", 0 } };
            double waterTotal = 0;
            
            try 
            {
                var dtWater = DataManager.GetTableData(DbName, "WaterMeterReadings", "日期", start, end);
                if (dtWater != null) 
                {
                    foreach (DataRow r in dtWater.Rows) 
                    {
                        if (dtWater.Columns.Contains("廢水處理量日統計") && double.TryParse(r["廢水處理量日統計"]?.ToString().Replace(",", ""), out double w)) waterTotal += w;
                    }
                }
            } 
            catch { }
            
            try 
            {
                var dtChem = DataManager.GetTableData(DbName, "WaterChemicals", "日期", start, end);
                if (dtChem != null) 
                {
                    foreach (DataRow r in dtChem.Rows) 
                    {
                        double pac = 0, naoh = 0, poly = 0;
                        if (dtChem.Columns.Contains("PACPAC_KG")) double.TryParse(r["PACPAC_KG"]?.ToString().Replace(",", ""), out pac); 
                        else if (dtChem.Columns.Contains("PAC_KG")) double.TryParse(r["PAC_KG"]?.ToString().Replace(",", ""), out pac); 
                        else if (dtChem.Columns.Contains("PAC日統計")) double.TryParse(r["PAC日統計"]?.ToString().Replace(",", ""), out pac);
                        
                        if (dtChem.Columns.Contains("NAOH_KG")) double.TryParse(r["NAOH_KG"]?.ToString().Replace(",", ""), out naoh); 
                        else if (dtChem.Columns.Contains("NAOH日統計")) double.TryParse(r["NAOH日統計"]?.ToString().Replace(",", ""), out naoh);
                        
                        if (dtChem.Columns.Contains("高分子_KG")) double.TryParse(r["高分子_KG"]?.ToString().Replace(",", ""), out poly); 
                        else if (dtChem.Columns.Contains("高分子日統計")) double.TryParse(r["高分子日統計"]?.ToString().Replace(",", ""), out poly);
                        
                        dict["PAC"] += pac; 
                        dict["NAOH"] += naoh; 
                        dict["高分子"] += poly;
                    }
                }
            } 
            catch { }
            
            if (waterTotal > 0) 
            { 
                dict["PAC/M3"] = dict["PAC"] / waterTotal; 
                dict["NAOH/M3"] = dict["NAOH"] / waterTotal; 
                dict["高分子/M3"] = dict["高分子"] / waterTotal; 
            }
            
            return dict;
        }

        private void FillDataPanels(string moduleName, Panel p1, Panel p2, Panel p3, Panel p4, Dictionary<string, double> curr, Dictionary<string, double> ly, Dictionary<string, double> l2y, bool isRecycleRate = false)
        {
            p1.Controls.Clear(); p2.Controls.Clear(); p3.Controls.Clear(); p4.Controls.Clear();

            bool dailyUsageHeaderAdded1 = false, dailyUsageHeaderAdded2 = false, monthlyHeaderAdded1 = false, monthlyHeaderAdded2 = false;

            foreach (var kvp in curr)
            {
                string key = kvp.Key;

                if (!IsVisible(moduleName, key)) continue;

                if (moduleName == "ChemStat") 
                {
                    if (key == "PAC") AddSectionHeader(p1, p2, p3, p4, "【藥劑使用量】");
                    else if (key == "PAC/M3") AddSectionHeader(p1, p2, p3, p4, "【每噸水藥劑量】");
                } 
                else if (moduleName == "DischargeStat") 
                {
                    if (key == "SS平均值") AddSectionHeader(p1, p2, p3, p4, "【水質】");
                } 
                else if (moduleName == "DailyUsage") 
                {
                    if (!dailyUsageHeaderAdded1 && (key == "廠區" || key == "行政區" || key == "全廠用量")) 
                    {
                        AddSectionHeader(p1, p2, p3, p4, "【自來水用量】"); 
                        dailyUsageHeaderAdded1 = true;
                    } 
                    else if (!dailyUsageHeaderAdded2 && (key.Contains("平均") || key.Contains("貯存區") || key.Contains("清水池"))) 
                    {
                        AddSectionHeader(p1, p2, p3, p4, "【平均使用量】"); 
                        dailyUsageHeaderAdded2 = true;
                    }
                } 
                else if (moduleName == "MonthlyVolume") 
                {
                    if (!monthlyHeaderAdded1 && (key == "廠區" || key == "行政區" || key == "全廠用量")) 
                    { 
                        AddSectionHeader(p1, p2, p3, p4, "【自來水量】"); 
                        monthlyHeaderAdded1 = true; 
                    } 
                    else if (!monthlyHeaderAdded2 && (key.Contains("月平均"))) 
                    { 
                        AddSectionHeader(p1, p2, p3, p4, "【平均使用量】"); 
                        monthlyHeaderAdded2 = true; 
                    }
                }

                double vCurr = kvp.Value;
                double vLy = ly.ContainsKey(key) ? ly[key] : 0;
                double vL2y = l2y.ContainsKey(key) ? l2y[key] : 0;

                p1.Controls.Add(CreateStatLabel(key, vCurr));
                p2.Controls.Add(CreateStatLabel(key, vLy));
                p3.Controls.Add(CreateStatLabel(key, vL2y));

                string diffText = "無基期"; 
                Color diffColor = Color.DimGray;
                
                if (vLy > 0 || (vCurr == 0 && vLy > 0)) 
                { 
                    double yoy = vLy == 0 ? 0 : ((vCurr - vLy) / vLy) * 100;
                    if (isRecycleRate && key.Contains("回收率")) yoy = vCurr - vLy; 

                    string formatStr = (key.Contains("回收率") || key.Contains("/M3") || key.Contains("污泥量") || key.Contains("平均") || key.Contains("最大") || key.Contains("最小")) ? "N1" : "N0";
                    diffText = (yoy > 0 ? "+" : "") + yoy.ToString(formatStr) + " %";
                    diffColor = yoy > 0 ? Color.IndianRed : (yoy < 0 ? Color.ForestGreen : Color.DimGray); 
                } 
                else if (vCurr > 0) 
                {
                    diffText = "新數據"; 
                    diffColor = Color.SteelBlue;
                }
                p4.Controls.Add(new Label { Text = $"{key}: {diffText}", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = diffColor, AutoSize = true, Margin = new Padding(0, 0, 0, 8) });
            }
        }

        private void AddSectionHeader(Panel p1, Panel p2, Panel p3, Panel p4, string title) 
        {
            p1.Controls.Add(CreateSectionHeader(title)); 
            p2.Controls.Add(CreateSectionHeader(title));
            p3.Controls.Add(CreateSectionHeader(title)); 
            p4.Controls.Add(CreateSectionHeader(title));
        }

        private Label CreateSectionHeader(string title) 
        { 
            return new Label { Text = title, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 10, 0, 5) }; 
        }

        private Label CreateStatLabel(string title, double value)
        {
            string unit = " M3", format = "N0"; 
            
            if (CustomUnitsCache.ContainsKey(title)) 
            {
                unit = " " + CustomUnitsCache[title];
                format = "N2"; 
            }
            else 
            {
                if (title.Contains("用電")) { unit = " KWH"; } 
                else if (title.Contains("%") || title.Contains("率")) { unit = " %"; format = "N1"; } 
                else if (title.Contains("包")) { unit = " 包"; } 
                else if (title.Contains("污泥量")) { unit = " 噸"; format = "N3"; } 
                else if (title.Contains("/M3")) { unit = " KG/M3"; format = "N3"; } 
                else if (title == "PAC" || title == "NAOH" || title == "高分子") { unit = " KG"; format = "N0"; }
                else if (title.Contains("月平均")) { unit = " M3/月"; format = "N1"; } 
                else if (title.Contains("平均") || title.Contains("最大") || title.Contains("最小")) 
                { 
                    if (title.Contains("SS") || title.Contains("COD") || title.Contains("BOD") || title.Contains("氨氮")) { unit = " mg/L"; format = "N2"; } 
                    else { unit = " M3/日"; format = "N1"; } 
                } 
            }

            return new Label { Text = $"{title}: {value.ToString(format)}{unit}", Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.FromArgb(45,45,45), AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        }

        private List<Panel> GetSelectedExportPanels()
        {
            List<Panel> selectedPanels = new List<Panel>();
            using (Form f = new Form() { Width = 400, Height = 400, Text = "選擇匯出項目", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
            {
                Label lbl = new Label { Text = "請勾選欲匯出至 PDF 的報表項目：", Dock = DockStyle.Top, Padding = new Padding(15, 15, 10, 5), Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                CheckedListBox clb = new CheckedListBox { Dock = DockStyle.Top, Height = 230, CheckOnClick = true, Font = new Font("Microsoft JhengHei UI", 14F), Margin = new Padding(10), BorderStyle = BorderStyle.None, BackColor = f.BackColor };
                clb.Items.Add("水資源數據統計", true); 
                clb.Items.Add("回收水統計", true); 
                clb.Items.Add("藥劑數據統計", true);
                clb.Items.Add("自來水用量統計", true); 
                clb.Items.Add("納管排放數據統計", true); 
                clb.Items.Add("自來水用量(繳費單)月統計", true);
                
                f.Controls.Add(clb);

                Button btnOk = new Button { Text = "確認匯出", Dock = DockStyle.Bottom, Height = 50, DialogResult = DialogResult.OK, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Cursor = Cursors.Hand };
                f.Controls.Add(btnOk);

                if (f.ShowDialog() == DialogResult.OK) 
                {
                    if (clb.GetItemChecked(0)) selectedPanels.Add(_pnlWaterBox);
                    if (clb.GetItemChecked(1)) selectedPanels.Add(_pnlRecycleBox);
                    if (clb.GetItemChecked(2)) selectedPanels.Add(_pnlChemBox);
                    if (clb.GetItemChecked(3)) selectedPanels.Add(_pnlDailyUsageBox);
                    if (clb.GetItemChecked(4)) selectedPanels.Add(_pnlDischargeBox);
                    if (clb.GetItemChecked(5)) selectedPanels.Add(_pnlMonthlyVolumeBox);
                }
            }
            return selectedPanels;
        }

        private void BtnPdf_Click(object sender, EventArgs e)
        {
            var panelsToExport = GetSelectedExportPanels();
            if (panelsToExport.Count == 0) return;

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF 檔案 (*.pdf)|*.pdf", FileName = "水資源綜合統計表_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                        List<Bitmap> bitmaps = new List<Bitmap>();
                        foreach (var pnl in panelsToExport) 
                        {
                            Bitmap bmp = new Bitmap(pnl.Width, pnl.Height);
                            pnl.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                            bitmaps.Add(bmp);
                        }

                        PrintDocument pd = new PrintDocument();
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                        pd.PrinterSettings.PrintToFile = true;
                        pd.PrinterSettings.PrintFileName = sfd.FileName;
                        pd.DefaultPageSettings.Landscape = true; 
                        pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);

                        int currentBmpIndex = 0;

                        pd.PrintPage += (s, ev) => 
                        {
                            Graphics g = ev.Graphics;
                            string headerText = $"導出日期：{DateTime.Now:yyyy-MM-dd HH:mm}   |   日報查詢：{_cboStartYear.Text}/{_cboStartMonth.Text}/{_cboStartDay.Text}~{_cboEndYear.Text}/{_cboEndMonth.Text}/{_cboEndDay.Text}   |   月報查詢：{_cboStartMonthYear.Text}/{_cboStartMonthMonth.Text}~{_cboEndMonthYear.Text}/{_cboEndMonthMonth.Text}";
                            g.DrawString(headerText, new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Brushes.Black, ev.MarginBounds.Left, ev.MarginBounds.Top - 15);

                            int currentY = ev.MarginBounds.Top + 15;
                            int bottomLimit = ev.MarginBounds.Bottom;

                            while (currentBmpIndex < bitmaps.Count) 
                            {
                                Bitmap bmp = bitmaps[currentBmpIndex];
                                float scale = (float)ev.MarginBounds.Width / bmp.Width;
                                int scaledHeight = (int)(bmp.Height * scale);

                                if (currentY + scaledHeight > bottomLimit) 
                                {
                                    if (currentY == ev.MarginBounds.Top + 15) 
                                    {
                                        scale = Math.Min(scale, (float)(bottomLimit - currentY) / bmp.Height);
                                        scaledHeight = (int)(bmp.Height * scale);
                                        g.DrawImage(bmp, ev.MarginBounds.Left, currentY, (int)(bmp.Width * scale), scaledHeight);
                                        currentY += scaledHeight + 20;
                                        currentBmpIndex++;
                                    } 
                                    else 
                                    {
                                        ev.HasMorePages = true;
                                        return;
                                    }
                                } 
                                else 
                                {
                                    g.DrawImage(bmp, ev.MarginBounds.Left, currentY, ev.MarginBounds.Width, scaledHeight);
                                    currentY += scaledHeight + 20; 
                                    currentBmpIndex++;
                                }
                            }
                            ev.HasMorePages = false;
                        };

                        pd.Print();
                        foreach (var bmp in bitmaps) bmp.Dispose();

                        MessageBox.Show("PDF 匯出成功！已根據項目自動分頁並全版面 A4 對齊。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("PDF 匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    } 
                    finally 
                    { 
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
                    }
                }
            }
        }
    }

    public static class CustomStatEngine
    {
        public class CustomRule 
        {
            public string Module { get; set; }
            public string StatName { get; set; }
            public string Unit { get; set; }
            public string Formula { get; set; }
        }

        private static List<CustomRule> _rules = new List<CustomRule>();
        private static readonly string ConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WaterCustomStats.txt");
        private static readonly string[] WaterTables = { "WaterMeterReadings", "WaterChemicals", "WaterUsageDaily", "DischargeData", "WaterVolume" };

        public static void LoadSettings()
        {
            _rules.Clear();
            App_WaterDashboard.CustomUnitsCache.Clear();
            if (File.Exists(ConfigFile)) 
            {
                try 
                {
                    foreach (var line in File.ReadAllLines(ConfigFile, Encoding.UTF8)) 
                    {
                        var p = line.Split('|');
                        if (p.Length >= 4) 
                        {
                            _rules.Add(new CustomRule { Module = p[0], StatName = p[1], Unit = p[2], Formula = p[3] });
                            App_WaterDashboard.CustomUnitsCache[p[1]] = p[2];
                        }
                    }
                } 
                catch { }
            }
        }

        public static void SaveSettings()
        {
            try 
            {
                var lines = _rules.Select(r => $"{r.Module}|{r.StatName}|{r.Unit}|{r.Formula}").ToArray();
                File.WriteAllLines(ConfigFile, lines, Encoding.UTF8);
            } 
            catch { }
        }

        public static List<CustomRule> GetStatsForModule(string moduleName)
        {
            return _rules.Where(r => r.Module == moduleName).ToList();
        }

        public static Dictionary<string, double> EvaluateForModule(string module, string startDate, string endDate)
        {
            var result = new Dictionary<string, double>();
            var rules = GetStatsForModule(module);
            if (rules.Count == 0) return result;

            foreach (var rule in rules)
            {
                string formula = rule.Formula;
                
                Regex regex = new Regex(@"(?<agg>SUM|AVG|MAX|MIN|COUNT)\(\[(?<table>[^\]]+)\]\.\[(?<col>[^\]]+)\]\)");
                var matches = regex.Matches(formula);

                foreach (Match m in matches)
                {
                    string agg = m.Groups["agg"].Value;
                    string table = m.Groups["table"].Value;
                    string col = m.Groups["col"].Value;
                    
                    // 🟢 動態判斷日期欄位：水污月表已改為 年月
                    string dateCol = (table == "DischargeData" || table == "WaterVolume") ? "年月" : "日期";

                    double computedValue = 0;
                    try 
                    {
                        DataTable dt = DataManager.GetTableData("Water", table, dateCol, startDate, endDate);
                        if (dt != null && dt.Columns.Contains(col)) 
                        {
                            List<double> values = new List<double>();
                            foreach (DataRow r in dt.Rows) 
                            {
                                if (double.TryParse(r[col]?.ToString().Replace(",", ""), out double v)) 
                                {
                                    values.Add(v);
                                }
                            }
                            
                            if (values.Count > 0 || agg == "COUNT") 
                            { 
                                if (agg == "SUM") computedValue = values.Sum();
                                else if (agg == "AVG") computedValue = values.Average();
                                else if (agg == "MAX") computedValue = values.Max();
                                else if (agg == "MIN") computedValue = values.Min();
                                else if (agg == "COUNT") computedValue = values.Count;
                            }
                        }
                    } 
                    catch { }

                    formula = formula.Replace(m.Value, computedValue.ToString());
                }

                double finalVal = 0;
                try 
                {
                    DataTable dtMath = new DataTable();
                    object computeResult = dtMath.Compute(formula, null);
                    if (computeResult != DBNull.Value) finalVal = Convert.ToDouble(computeResult);
                } 
                catch 
                { 
                    finalVal = 0; 
                }

                result[rule.StatName] = finalVal;
            }

            return result;
        }

        public static void OpenConfigDialog(string moduleName)
        {
            using (Form f = new Form() { Width = 1050, Height = 650, Text = $"統計設定引擎 ({moduleName})", StartPosition = FormStartPosition.CenterParent, Font = new Font("Microsoft JhengHei UI", 12F) })
            {
                Label lblTop = new Label { Text = "📝 自定義延伸統計公式", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Dock = DockStyle.Top, Padding = new Padding(10), AutoSize = true };

                Panel leftPnl = new Panel { Dock = DockStyle.Left, Width = 280, Padding = new Padding(10) };
                ListBox lbExisting = new ListBox { Dock = DockStyle.Fill };
                var existingRules = GetStatsForModule(moduleName);
                
                foreach (var r in existingRules) lbExisting.Items.Add(r.StatName);
                
                leftPnl.Controls.Add(lbExisting);
                leftPnl.Controls.Add(new Label { Text = "已建立的統計項目:", Dock = DockStyle.Top, Height = 30 });

                Panel rightPnl = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };

                int comboW = 180;
                int btnW = 120;

                FlowLayoutPanel pnlTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 0, 0, 15) };
                TextBox txtName = new TextBox { Width = 250 };
                TextBox txtUnit = new TextBox { Width = 100 };
                
                pnlTop.Controls.AddRange(new Control[] {
                    new Label { Text = "顯示名稱:", AutoSize = true, Margin = new Padding(0, 5, 5, 0) }, txtName,
                    new Label { Text = "單位:", AutoSize = true, Margin = new Padding(20, 5, 5, 0) }, txtUnit
                });

                GroupBox boxBuilder = new GroupBox { Text = "建立來源欄位", Dock = DockStyle.Top, Height = 100, Padding = new Padding(10) };
                FlowLayoutPanel pnlFields = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
                
                ComboBox cbTables = new ComboBox { Width = comboW, DropDownStyle = ComboBoxStyle.DropDownList };
                ComboBox cbCols = new ComboBox { Width = comboW, DropDownStyle = ComboBoxStyle.DropDownList };
                ComboBox cbAggs = new ComboBox { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
                
                cbAggs.Items.AddRange(new string[] { "SUM", "AVG", "MAX", "MIN", "COUNT" }); 
                cbAggs.SelectedIndex = 0;
                
                Button btnInsert = new Button { Text = "插入公式 ⬇️", Width = btnW, Height = 32, BackColor = Color.LightBlue };

                cbTables.Items.AddRange(WaterTables);
                cbTables.SelectedIndexChanged += (s, e) => {
                    cbCols.Items.Clear();
                    if (cbTables.SelectedItem != null) 
                    {
                        var cols = DataManager.GetColumnNames("Water", cbTables.SelectedItem.ToString());
                        // 🟢 過濾時同步排除 年月
                        foreach (var c in cols) 
                        {
                            if (c != "Id" && c != "日期" && c != "月份" && c != "年月" && c != "備註") 
                            {
                                cbCols.Items.Add(c);
                            }
                        }
                        if (cbCols.Items.Count > 0) cbCols.SelectedIndex = 0;
                    }
                };

                pnlFields.Controls.AddRange(new Control[] { cbTables, cbCols, cbAggs, btnInsert });
                boxBuilder.Controls.Add(pnlFields);

                FlowLayoutPanel pnlKeys = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(0, 10, 0, 5) };
                string[] keys = { "+", "-", "*", "/", "(", ")" };
                
                foreach (var k in keys) 
                {
                    Button b = new Button { Text = k, Width = 50, Height = 35 };
                    pnlKeys.Controls.Add(b);
                }

                RichTextBox rtbFormula = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 14F), BackColor = Color.AliceBlue };
                Label lblF = new Label { Text = "計算公式: (支援數學運算子與欄位變數)", Dock = DockStyle.Top, Height = 30, Margin = new Padding(0,10,0,0) };

                foreach (Control c in pnlKeys.Controls) 
                {
                    if (c is Button b) b.Click += (s, e) => rtbFormula.AppendText(" " + b.Text + " ");
                }
                
                btnInsert.Click += (s, e) => {
                    if (cbTables.SelectedItem != null && cbCols.SelectedItem != null && cbAggs.SelectedItem != null)
                        rtbFormula.AppendText($"{cbAggs.Text}([{cbTables.Text}].[{cbCols.Text}])");
                };

                rightPnl.Controls.Add(rtbFormula);
                rightPnl.Controls.Add(lblF);
                rightPnl.Controls.Add(pnlKeys);
                rightPnl.Controls.Add(boxBuilder);
                rightPnl.Controls.Add(pnlTop);

                Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(10) };
                Button btnSave = new Button { Text = "💾 儲存並套用", Width = 150, Dock = DockStyle.Right, BackColor = Color.ForestGreen, ForeColor = Color.White };
                Button btnDel = new Button { Text = "🗑️ 刪除", Width = 100, Dock = DockStyle.Left, BackColor = Color.IndianRed, ForeColor = Color.White };

                lbExisting.SelectedIndexChanged += (s, e) => {
                    if (lbExisting.SelectedIndex >= 0) 
                    {
                        var rule = _rules.First(r => r.Module == moduleName && r.StatName == lbExisting.SelectedItem.ToString());
                        txtName.Text = rule.StatName;
                        txtUnit.Text = rule.Unit;
                        rtbFormula.Text = rule.Formula;
                    }
                };

                btnSave.Click += (s, e) => {
                    if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(rtbFormula.Text)) 
                    {
                        MessageBox.Show("名稱與公式不可空白！"); 
                        return;
                    }
                    _rules.RemoveAll(r => r.Module == moduleName && r.StatName == txtName.Text);
                    _rules.Add(new CustomRule { Module = moduleName, StatName = txtName.Text, Unit = txtUnit.Text, Formula = rtbFormula.Text });
                    SaveSettings(); 
                    LoadSettings();
                    f.DialogResult = DialogResult.OK;
                };

                btnDel.Click += (s, e) => {
                    if (lbExisting.SelectedItem != null) 
                    {
                        _rules.RemoveAll(r => r.Module == moduleName && r.StatName == lbExisting.SelectedItem.ToString());
                        SaveSettings(); 
                        LoadSettings();
                        f.DialogResult = DialogResult.OK;
                    }
                };

                pnlBottom.Controls.Add(btnSave);
                pnlBottom.Controls.Add(btnDel);

                f.Controls.Add(rightPnl);
                f.Controls.Add(leftPnl);
                f.Controls.Add(pnlBottom);
                f.Controls.Add(lblTop);

                f.ShowDialog();
            }
        }
    }
}
