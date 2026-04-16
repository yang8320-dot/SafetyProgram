/// FILE: Safety_System/App_GenericTable.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml; 

namespace Safety_System
{
    public class App_GenericTable
    {
        private enum TimeMode { Date, YearMonth, Year }
        private TimeMode _timeMode = TimeMode.Date;

        private DataGridView _dgv;
        private ComboBox _cboStartYear, _cboStartMonth, _cboStartDay;
        private ComboBox _cboEndYear, _cboEndMonth, _cboEndDay;
        
        private TextBox _txtNewColName, _txtRenameCol;
        private ComboBox _cboColumns;
        private GroupBox _boxAdvanced; 
        
        private Button _btnToggle, _btnRead, _btnSave, _btnExport, _btnImport;
        private Label _lblStatus;     

        private bool _isFirstLoad = true;
        
        private readonly string _dbName; 
        private readonly string _tableName; 
        private readonly string _chineseTitle;
        
        private string _dateColumnName = "日期";

        private DataGridViewAutoCalcHelper _calcHelper; 

        private readonly Dictionary<string, string> _schemaMap = new Dictionary<string, string>
        {
            { "TargetManagement", "[年度] TEXT, [修訂日] TEXT, [單位] TEXT, [目標名稱] TEXT, [管理目標計畫表編號] TEXT, [施實重點項目1] TEXT, [日程1] TEXT, [施實重點項目2] TEXT, [日程2] TEXT, [施實重點項目3] TEXT, [日程3] TEXT, [施實重點項目4] TEXT, [日程4] TEXT, [施實重點項目5] TEXT, [日程5] TEXT, [預估成本] TEXT, [預估成效] TEXT, [計畫績效指標] TEXT, [績效指標計算方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "NearMiss", "[日期] TEXT, [時間] TEXT, [文件編號] TEXT, [單位] TEXT, [地點] TEXT, [事故類別] TEXT, [事故類型] TEXT, [發生經過] TEXT, [改善對策] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SafetyInspection", "[日期] TEXT, [巡檢區域] TEXT, [檢查項目] TEXT, [檢查結果] TEXT, [缺失描述] TEXT, [改善措施] TEXT, [負責人] TEXT, [狀態] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SafetyObservation", "[日期] TEXT, [區域] TEXT, [類別] TEXT, [描述] TEXT, [觀查人] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "TrafficInjury", "[日期] TEXT, [姓名] TEXT, [地點] TEXT, [狀態] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkInjury", "[日期] TEXT, [姓名] TEXT, [受傷部位] TEXT, [原因] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "MinorInjury", "[日期] TEXT, [姓名] TEXT, [單位] TEXT, [發生經過] TEXT, [處置方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_IL", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [丁基膠MT] TEXT, [結構膠MT] TEXT, [鋁條MT] TEXT, [乾燥劑MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_LM", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [PVB膜MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_CR", "[年月] TEXT, [鍍膜加不良MT] TEXT, [鍍膜生產量MT] TEXT, [除膜生產加不良MT] TEXT, [除膜成品產量MT] TEXT, [靶材MT] TEXT, [隔離粉MT] TEXT, [氧化鈰MT] TEXT, [噴砂底板成品鋁MT] TEXT, [噴砂底板成品其他MT] TEXT, [氧化鋁砂金鋼砂MT] TEXT, [D1099MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_T", "[年月] TEXT, [強化生產加不良MT] TEXT, [強化生產量] TEXT, [砂布輪MT] TEXT, [砂帶MT] TEXT, [印刷生產加不良MT] TEXT, [印刷生產量MT] TEXT, [油墨MT] TEXT, [網板清洗劑MT] TEXT, [D1599MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_GCTE", "[年月] TEXT, [切片生產加不良MT] TEXT, [切片生產量MT] TEXT, [磨邊生產量不良MT] TEXT, [砂布輪砂輪片MT] TEXT, [鑽孔生產加不良MT] TEXT, [鑽孔生產量MT] TEXT, [] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_ML", "[年月] TEXT, [甲醇MT] TEXT, [乙醇MT] TEXT, [潤滑油MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_Water", "[年月] TEXT, [D0902MT] TEXT, [D0201MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_Performance", "[年月] TEXT, [單位] TEXT, [項目] TEXT, [說明] TEXT, [預計執行週期] TEXT, [預估可節省或改善之數據] TEXT, [費用TWD] TEXT, [回應窗口] TEXT, [績效追蹤1] TEXT, [績效追蹤2] TEXT, [統計至12月底之實際數據含計算式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "AirPollution", "[年度] TEXT, [季度] TEXT, [排放量] TEXT, [繳費金額] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "HazardStats", "[年月] TEXT, [品項] TEXT, [單位] TEXT, [數量] TEXT, [使用量] TEXT, [庫存量] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkInjuryReport", "[年月] TEXT, [受僱勞工男性人數] TEXT, [非屬受僱勞工之其他工作者男性人數] TEXT, [受僱勞工女性人數] TEXT, [非屬受僱勞工之其他工作者女性人數] TEXT, [受僱勞工總計工作日數] TEXT, [非屬受僱勞工之其他工作者總計工作日數] TEXT, [受僱勞工總經歷工時] TEXT, [非屬受僱勞工之其他工作者總經歷工時] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ChemRegulations", "[類別] TEXT, [公告日期] TEXT, [法規名稱] TEXT, [法規內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CAS No] TEXT, [檢測週期] TEXT, [容許濃度ppm] TEXT, [容許濃度mgm3] TEXT, [管制濃度] TEXT, [分級運作量] TEXT, [管制行為] TEXT, [定期申報頻率] TEXT, [毒性分類] TEXT, [記錄] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SDS_Inventory", "[日期] TEXT, [廠內編號] TEXT, [化學品名稱] TEXT, [CAS_No] TEXT, [危害成份] TEXT, [危害分類] TEXT, [供應商] TEXT, [SDS版本日期] TEXT, [存放地點] TEXT, [最大儲存量] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FireResponsible", "[單位] TEXT, [場所區域] TEXT, [防火負責人] TEXT, [防火負責人職稱] TEXT, [火源責任人] TEXT, [火源責任人職稱] TEXT, [責任代理人] TEXT, [責任代理人職稱] TEXT, [更新日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FireEquip", "[日期] TEXT, [設備名稱] TEXT, [編號] TEXT, [位置] TEXT, [有效日期] TEXT, [檢查結果] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "訓練時數", "[日期] TEXT, [員工姓名] TEXT, [受訓項目] TEXT, [課程名稱] TEXT, [訓練時數] TEXT, [HR外訓申請] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "EnvMonitor", "[日期] TEXT, [SEG編號] TEXT, [測點名稱] TEXT, [噪音_db] TEXT, [粉塵_區域] TEXT, [粉塵_個人] TEXT, [一氧化鉛] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastewaterPeriodic", "[日期] TEXT, [申報季別] TEXT, [排放水量] TEXT, [COD] TEXT, [SS] TEXT, [BOD] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DrinkingWater", "[日期] TEXT, [採樣點位置] TEXT, [大腸桿菌群] TEXT, [總菌落數] TEXT, [鉛] TEXT, [濁度] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "IndustrialZoneTest", "[日期] TEXT, [採樣點位置] TEXT, [水溫] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [重金屬] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SoilGasTest", "[日期] TEXT, [採樣井編號] TEXT, [測漏氣體濃度] TEXT, [甲烷] TEXT, [二氧化碳] TEXT, [氧氣] TEXT, [檢測機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastewaterSelfTest", "[日期] TEXT, [採樣時間] TEXT, [採樣位置] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [透視度] TEXT, [檢驗人員] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "CoolingWaterVendor", "[日期] TEXT, [廠商名稱] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [添加藥劑] TEXT, [檢驗結果] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "CoolingWaterSelf", "[日期] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [檢驗人員] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "TCLP", "[日期] TEXT, [樣品名稱] TEXT, [總鉛] TEXT, [鏓鉻] TEXT, [鏓銅] TEXT, [鏓鋇] TEXT, [總鎘] TEXT, [總硒] TEXT, [六價鉻] TEXT, [總汞] TEXT, [總砷] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterMeterCalibration", "[日期] TEXT, [水錶編號] TEXT, [水錶位置] TEXT, [校正前讀數] TEXT, [校正後讀數] TEXT, [校正單位] TEXT, [下次校正日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "OtherTests", "[日期] TEXT, [檢測項目] TEXT, [檢測位置] TEXT, [檢測數值] TEXT, [單位] TEXT, [合格標準] TEXT, [檢測機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "HealthPromotion", "[日期] TEXT, [活動名稱] TEXT, [參與人數] TEXT, [執行單位] TEXT, [成果摘要] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			{ "fire_check_stats", "[日期] TEXT, [年月] TEXT, [單位] TEXT, [日常火源] TEXT, [日常火源備註] TEXT, [消防設備] TEXT, [消防設備備註] TEXT, [防火避難設施] TEXT, [防火避難設施備註] TEXT, [附件檔案] TEXT, [備註] TEXT" }
        };

        public App_GenericTable(string dbName, string tableName, string chineseTitle)
        {
            _dbName = dbName;
            _tableName = tableName;
            _chineseTitle = chineseTitle;
        }

        private string GetExpectedFolderName(string rowDateStr)
        {
            if (string.IsNullOrWhiteSpace(rowDateStr)) return DateTime.Now.ToString("yyyy-MM");
            
            if (_timeMode == TimeMode.Year) 
            {
                if (rowDateStr.Length >= 4) return rowDateStr.Substring(0, 4);
            }
            else 
            {
                if (rowDateStr.Length >= 7) return rowDateStr.Substring(0, 7);
            }
            return DateTime.Now.ToString("yyyy-MM");
        }

        public Control GetView()
        {
            string schema = _schemaMap.ContainsKey(_tableName) ? _schemaMap[_tableName] : "[日期] TEXT, [備註] TEXT";
            string createSql = $"CREATE TABLE IF NOT EXISTS [{_tableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
            DataManager.InitTable(_dbName, _tableName, createSql);

            List<string> columns = DataManager.GetColumnNames(_dbName, _tableName);
            if (columns.Contains("月份")) 
            {
                try 
                {
                    DataManager.RenameColumn(_dbName, _tableName, "月份", "年月");
                    columns = DataManager.GetColumnNames(_dbName, _tableName);
                } 
                catch { }
            }

            if (columns.Contains("日期")) 
            { 
                _timeMode = TimeMode.Date; 
                _dateColumnName = "日期"; 
            }
            else if (columns.Contains("年月")) 
            { 
                _timeMode = TimeMode.YearMonth; 
                _dateColumnName = "年月"; 
            }
            else if (columns.Contains("年度")) 
            { 
                _timeMode = TimeMode.Year; 
                _dateColumnName = "年度"; 
            }
            else 
            {
                string altDate = columns.FirstOrDefault(c => c.Contains("日期"));
                if (altDate != null) 
                { 
                    _timeMode = TimeMode.Date; 
                    _dateColumnName = altDate; 
                }
                else 
                {
                    altDate = columns.FirstOrDefault(c => c.Contains("年月"));
                    if (altDate != null) 
                    { 
                        _timeMode = TimeMode.YearMonth; 
                        _dateColumnName = altDate; 
                    }
                    else 
                    {
                        altDate = columns.FirstOrDefault(c => c.Contains("年度"));
                        if (altDate != null) 
                        { 
                            _timeMode = TimeMode.Year; 
                            _dateColumnName = altDate; 
                        }
                        else 
                        { 
                            _timeMode = TimeMode.Date; 
                            _dateColumnName = columns.FirstOrDefault(c => c != "Id") ?? "Id"; 
                        }
                    }
                }
            }

            TableLayoutPanel main = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, Padding = new Padding(15) };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

            GroupBox boxTop = new GroupBox { Text = $"{_chineseTitle} (庫：{_dbName} 表：{_tableName})", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), AutoSize = true, Padding = new Padding(10, 15, 10, 10), Margin = new Padding(0, 0, 0, 10) };
            FlowLayoutPanel row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            
            Label lblRange = new Label { Text = "查詢區間:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            
            _cboStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboStartDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndMonth = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboEndDay = new ComboBox { Width = 55, DropDownStyle = ComboBoxStyle.DropDownList };

            int currentYear = DateTime.Now.Year;
            for (int i = currentYear - 25; i <= currentYear + 25; i++) {
                _cboStartYear.Items.Add(i); 
                _cboEndYear.Items.Add(i);
            }
            for (int i = 1; i <= 12; i++) {
                _cboStartMonth.Items.Add(i.ToString("D2")); 
                _cboEndMonth.Items.Add(i.ToString("D2"));
            }
            for (int i = 1; i <= 31; i++) {
                _cboStartDay.Items.Add(i.ToString("D2")); 
                _cboEndDay.Items.Add(i.ToString("D2"));
            }

            if (_timeMode == TimeMode.YearMonth || _timeMode == TimeMode.Year) 
            {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddMonths(-6));
            }
            else 
            {
                SetComboDate(_cboStartYear, _cboStartMonth, _cboStartDay, DateTime.Today.AddDays(-30));
            }
            
            SetComboDate(_cboEndYear, _cboEndMonth, _cboEndDay, DateTime.Today);

            _btnRead = new Button { Text = "🔍 讀取資料", Size = new Size(130, 35), BackColor = Color.WhiteSmoke };
            _btnRead.Click += async (s, e) => { _isFirstLoad = false; await LoadGridDataAsync(); };

            _btnSave = new Button { Name = "btnSave", Text = "💾 儲存數據", Size = new Size(130, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            _btnSave.Click += BtnSave_Click; 
            
            _btnExport = new Button { Text = "📤 匯出Excel", Size = new Size(130, 35) }; 
            _btnExport.Click += BtnExport_Click;

            _btnImport = new Button { Text = "📥 匯入Excel", Size = new Size(130, 35) }; 
            _btnImport.Click += BtnImportExcel_Click;

            _btnToggle = new Button { Text = "[ + ] 進階管理", Size = new Size(130, 35), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
            _btnToggle.Click += (s, e) => {
                _boxAdvanced.Visible = !_boxAdvanced.Visible;
                _btnToggle.Text = _boxAdvanced.Visible ? "[ - ] 隱藏管理" : "[ + ] 進階管理";
            };

            Label lblSY = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblSM = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblSD = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 8, 5, 0) };
            Label lblEY = new Label { Text = "年", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblEM = new Label { Text = "月", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };
            Label lblED = new Label { Text = "日", AutoSize = true, Margin = new Padding(0, 8, 5, 0) };

            if (_timeMode == TimeMode.YearMonth) 
            {
                _cboStartDay.Visible = false; lblSD.Visible = false;
                _cboEndDay.Visible = false; lblED.Visible = false;
            } 
            else if (_timeMode == TimeMode.Year) 
            {
                _cboStartDay.Visible = false; lblSD.Visible = false;
                _cboEndDay.Visible = false; lblED.Visible = false;
                _cboStartMonth.Visible = false; lblSM.Visible = false;
                _cboEndMonth.Visible = false; lblEM.Visible = false;
            }

            row1.Controls.AddRange(new Control[] {
                lblRange, _cboStartYear, lblSY, _cboStartMonth, lblSM, _cboStartDay, lblSD,
                lblTilde, _cboEndYear, lblEY, _cboEndMonth, lblEM, _cboEndDay, lblED,
                _btnRead, _btnExport, _btnImport, _btnToggle, _btnSave
            });

            boxTop.Controls.Add(row1);

            _boxAdvanced = new GroupBox { Text = "進階欄位與權限操作", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F), AutoSize = true, Visible = false, Padding = new Padding(10, 15, 10, 10), ForeColor = Color.DimGray, Margin = new Padding(0, 0, 0, 10) };
            FlowLayoutPanel flpAdv = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            
            FlowLayoutPanel rowAdv1 = new FlowLayoutPanel { AutoSize = true };
            _txtNewColName = new TextBox { Width = 150 };
            
            Button bAdd = new Button { Text = "新增欄位", Size = new Size(100, 35) };
            bAdd.Click += async (s, e) => { 
                if (!string.IsNullOrEmpty(_txtNewColName.Text) && AuthManager.VerifyAdmin()) 
                { 
                    DataManager.AddColumn(_dbName, _tableName, _txtNewColName.Text); 
                    await LoadGridDataAsync(); 
                    _txtNewColName.Clear(); 
                } 
            };
            
            _cboColumns = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList }; 
            _txtRenameCol = new TextBox { Width = 120 };
            
            Button bRen = new Button { Text = "修改名稱", Size = new Size(100, 35) };
            bRen.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && !string.IsNullOrEmpty(_txtRenameCol.Text) && AuthManager.VerifyAdmin()) 
                { 
                    DataManager.RenameColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString(), _txtRenameCol.Text); 
                    await LoadGridDataAsync(); 
                    _txtRenameCol.Clear(); 
                } 
            };
            
            Button bDelCol = new Button { Text = "刪除整欄", Size = new Size(100, 35), BackColor = Color.DarkOrange, ForeColor = Color.White };
            bDelCol.Click += async (s, e) => { 
                if (_cboColumns.SelectedItem != null && AuthManager.VerifyAdmin()) 
                { 
                    if(MessageBox.Show($"確定刪除整欄【{_cboColumns.SelectedItem}】？", "確認", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    { 
                        DataManager.DropColumn(_dbName, _tableName, _cboColumns.SelectedItem.ToString()); 
                        await LoadGridDataAsync(); 
                    } 
                } 
            };
            
            Button bDelRow = new Button { Text = "🗑 刪除選取列", Size = new Size(120, 35), BackColor = Color.IndianRed, ForeColor = Color.White };
            bDelRow.Click += async (s, e) => {
                var selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>()
                                       .Select(c => c.OwningRow)
                                       .Where(r => !r.IsNewRow && r.Cells["Id"].Value != DBNull.Value)
                                       .Distinct().ToList();
                                       
                if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(包含所屬的實體附件檔案也將被永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                {
                    if (AuthManager.VerifyUser()) 
                    {
                        foreach (var r in selectedRows) 
                        {
                            if (_dgv.Columns.Contains("附件檔案")) 
                            {
                                string relPathStr = r.Cells["附件檔案"].Value?.ToString();
                                if (!string.IsNullOrEmpty(relPathStr)) 
                                {
                                    string[] paths = relPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var p in paths) DeletePhysicalFile(p, r.Index);
                                }
                            }
                            DataManager.DeleteRecord(_dbName, _tableName, Convert.ToInt32(r.Cells["Id"].Value));
                        }
                        await LoadGridDataAsync(); 
                        MessageBox.Show("刪除成功！");
                    }
                }
            };

            rowAdv1.Controls.AddRange(new Control[] { new Label { Text = "欄位/列操作:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, _txtNewColName, bAdd, _cboColumns, _txtRenameCol, bRen, bDelCol, bDelRow });
            
            FlowLayoutPanel rowAdv2 = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            TextBox txtLimit = new TextBox { Width = 100, Text = "100" };
            Button bLimitRead = new Button { Text = "讀取指定筆數", Size = new Size(120, 35), BackColor = Color.SteelBlue, ForeColor = Color.White };
            bLimitRead.Click += async (s, e) => { 
                if (int.TryParse(txtLimit.Text, out int l)) 
                { 
                    SetUIState(false, "讀取中...", Color.Orange);
                    DataTable dt = null;
                    await Task.Run(() => {
                        dt = DataManager.GetLatestRecords(_dbName, _tableName, l); 
                        EnforceDateFormats(dt);
                    });
                    _dgv.DataSource = dt; 
                    ApplyGridStyles(); 
                    RestoreColumnOrder();
                    SetUIState(true, $"載入成功，共 {dt.Rows.Count} 筆", Color.Green);
                } 
            };
            
            rowAdv2.Controls.AddRange(new Control[] { new Label { Text = "調閱最近寫入筆數:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) }, txtLimit, bLimitRead });
            
            flpAdv.Controls.Add(rowAdv1); 
            flpAdv.Controls.Add(rowAdv2);
            _boxAdvanced.Controls.Add(flpAdv);

            _lblStatus = new Label { Text = "系統就緒", ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 5) };

            _dgv = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = true, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AllowUserToOrderColumns = true,
                Margin = new Padding(0, 10, 0, 10)
            };
            
            _dgv.RowTemplate.Height = 35;
            _dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            
            _dgv.CellFormatting += Dgv_CellFormatting;
            _dgv.CellClick += Dgv_CellClick;
            _dgv.KeyDown += Dgv_KeyDown;
            
            _calcHelper = new DataGridViewAutoCalcHelper(_dgv);

            main.Controls.Add(boxTop, 0, 0); 
            main.Controls.Add(_boxAdvanced, 0, 1); 
            main.Controls.Add(_lblStatus, 0, 2);
            main.Controls.Add(_dgv, 0, 3);

            _ = LoadGridDataAsync(); 
            return main;
        }

        private void SetUIState(bool isEnabled, string statusText, Color statusColor) 
        {
            _btnRead.Enabled = isEnabled; 
            _btnSave.Enabled = isEnabled; 
            _btnImport.Enabled = isEnabled; 
            _btnExport.Enabled = isEnabled;
            
            _lblStatus.Text = statusText; 
            _lblStatus.ForeColor = statusColor;
        }

        private void ApplyGridStyles() 
        {
            if (_dgv.Columns.Contains("Id")) 
            {
                _dgv.Columns["Id"].ReadOnly = true;
                _dgv.Columns["Id"].Visible = false;
            }
            
            if (_dgv.Columns.Contains(_dateColumnName)) 
            {
                string fmt = "yyyy-MM-dd";
                if (_timeMode == TimeMode.YearMonth) fmt = "yyyy-MM";
                else if (_timeMode == TimeMode.Year) fmt = "yyyy";
                
                _dgv.Columns[_dateColumnName].DefaultCellStyle.Format = fmt;
            }
            
            foreach (DataGridViewColumn col in _dgv.Columns) 
            {
                if (col.Name.Contains("附件檔案")) 
                {
                    col.ReadOnly = true; 
                    col.DefaultCellStyle.ForeColor = Color.Blue;
                    col.DefaultCellStyle.Font = new Font(_dgv.Font, FontStyle.Underline);
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }
        }

        private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) 
            {
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                if (colName.Contains("附件檔案") && e.Value != null) 
                {
                    string pathStr = e.Value.ToString();
                    if (!string.IsNullOrEmpty(pathStr)) 
                    {
                        string[] parts = pathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 1) 
                        {
                            e.Value = $"📁 [共 {parts.Length} 個檔案]";
                        } 
                        else 
                        {
                            e.Value = Path.GetFileName(parts[0]);
                        }
                        e.FormattingApplied = true;
                    }
                }
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e) 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < _dgv.Rows.Count && !_dgv.Rows[e.RowIndex].IsNewRow) 
            {
                if (_dgv.Columns[e.ColumnIndex].Name.Contains("附件檔案")) 
                {
                    string currentVal = _dgv[e.ColumnIndex, e.RowIndex].Value?.ToString();
                    
                    string rowDateStr = _dgv[_dateColumnName, e.RowIndex].Value?.ToString() ?? "";
                    string targetFolder = GetExpectedFolderName(rowDateStr);

                    using (var frm = new AttachmentForm(currentVal, _dbName, _tableName, targetFolder, path => DeletePhysicalFile(path, e.RowIndex))) 
                    {
                        if (frm.ShowDialog() == DialogResult.OK) 
                        {
                            _dgv[e.ColumnIndex, e.RowIndex].Value = frm.FinalPathsString;
                            _dgv.EndEdit();
                        }
                    }
                }
            }
        }

        private bool IsFileUsedInDatabase(string relativePath)
        {
            try 
            {
                DataTable dt = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                foreach (DataRow row in dt.Rows) 
                {
                    string val = row["附件檔案"]?.ToString();
                    if (!string.IsNullOrEmpty(val) && val.Contains(relativePath)) 
                    {
                        return true;
                    }
                }
                return false;
            } 
            catch { return true; } 
        }

        private void DeletePhysicalFile(string relativePath, int currentRowIndex) 
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return;
            
            bool isUsedByOthers = false;
            foreach (DataGridViewRow row in _dgv.Rows) 
            {
                if (row.Index == currentRowIndex || row.IsNewRow) continue;
                
                if (_dgv.Columns.Contains("附件檔案")) 
                {
                    string cellVal = row.Cells["附件檔案"].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellVal)) 
                    {
                        string[] paths = cellVal.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (paths.Contains(relativePath)) 
                        { 
                            isUsedByOthers = true; 
                            break; 
                        }
                    }
                }
            }
            
            if (!isUsedByOthers && IsFileUsedInDatabase(relativePath)) 
            {
                isUsedByOthers = true;
            }
            
            if (isUsedByOthers) return;
            
            try 
            {
                string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                if (File.Exists(absPath)) 
                {
                    File.Delete(absPath); 
                    DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(absPath));
                    string attachRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");
                    
                    while (dir != null && dir.FullName.StartsWith(attachRootDir) && dir.FullName.Length > attachRootDir.Length) 
                    {
                        if (dir.Exists && dir.GetFiles().Length == 0 && dir.GetDirectories().Length == 0) 
                        {
                            dir.Delete(); 
                            dir = dir.Parent;
                        } 
                        else 
                        { 
                            break; 
                        }
                    }
                }
            } 
            catch { }
        }

        private void EnforceDateFormats(DataTable dt) 
        {
            if (dt == null || !dt.Columns.Contains(_dateColumnName)) return;
            
            string format = "yyyy-MM-dd";
            if (_timeMode == TimeMode.YearMonth) format = "yyyy-MM";
            else if (_timeMode == TimeMode.Year) format = "yyyy";
            
            foreach (DataRow row in dt.Rows) 
            {
                if (row.RowState == DataRowState.Deleted) continue;
                
                string val = row[_dateColumnName]?.ToString();
                
                if (!string.IsNullOrWhiteSpace(val)) 
                {
                    val = val.Replace("/", "-");
                    if (_timeMode == TimeMode.Year && val.Length == 4 && int.TryParse(val, out _)) 
                    {
                        row[_dateColumnName] = val; 
                        continue;
                    }
                    if (DateTime.TryParse(val, out DateTime d)) 
                    {
                        row[_dateColumnName] = d.ToString(format);
                    }
                }
            }
        }

        private async Task LoadGridDataAsync() 
        {
            SetUIState(false, "資料庫讀取中，請稍候...", Color.Orange);
            DataTable dt = null;
            
            string sDate = GetDateString(_cboStartYear, _cboStartMonth, _cboStartDay);
            string eDate = GetDateString(_cboEndYear, _cboEndMonth, _cboEndDay);
            
            await Task.Run(() => {
                if (_isFirstLoad) 
                { 
                    dt = DataManager.GetLatestRecords(_dbName, _tableName, 30); 
                    _isFirstLoad = false; 
                } 
                else 
                { 
                    dt = DataManager.GetTableData(_dbName, _tableName, _dateColumnName, sDate, eDate); 
                }
                EnforceDateFormats(dt); 
            });
            
            _dgv.DataSource = dt; 
            ApplyGridStyles(); 
            UpdateCboColumns(); 
            RestoreColumnOrder();
            SetUIState(true, $"讀取成功，共載入 {dt.Rows.Count} 筆資料", Color.Green);
        }

        private string GetDateString(ComboBox y, ComboBox m, ComboBox d) 
        {
            if (_timeMode == TimeMode.Year) return y.SelectedItem.ToString();
            if (_timeMode == TimeMode.YearMonth) return $"{y.SelectedItem}-{m.SelectedItem}";
            
            return $"{y.SelectedItem}-{m.SelectedItem}-{d.SelectedItem}";
        }

        private void UpdateCboColumns() 
        {
            _cboColumns.Items.Clear();
            foreach (DataGridViewColumn c in _dgv.Columns) 
            {
                if (c.Name != "Id" && c.Name != _dateColumnName) 
                {
                    _cboColumns.Items.Add(c.Name);
                }
            }
        }

        private void SetComboDate(ComboBox y, ComboBox m, ComboBox d, DateTime date) 
        {
            if (y.Items.Contains(date.Year)) y.SelectedItem = date.Year;
            m.SelectedItem = date.Month.ToString("D2"); 
            d.SelectedItem = date.Day.ToString("D2");
        }

        private void SaveColumnOrder() 
        { 
            try 
            { 
                var ordered = _dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).Select(c => c.Name).ToArray(); 
                File.WriteAllText($"ColOrder_{_dbName}_{_tableName}.txt", string.Join(",", ordered), Encoding.UTF8); 
            } 
            catch { } 
        }
        
        private void RestoreColumnOrder() 
        { 
            try 
            { 
                string fn = $"ColOrder_{_dbName}_{_tableName}.txt"; 
                if (File.Exists(fn)) 
                { 
                    string[] saved = File.ReadAllText(fn, Encoding.UTF8).Split(','); 
                    for (int i = 0; i < saved.Length; i++) 
                    {
                        if (_dgv.Columns.Contains(saved[i])) _dgv.Columns[saved[i]].DisplayIndex = i; 
                    }
                } 
            } 
            catch { } 
        }

        private void SyncAttachmentPaths(DataTable dt) 
        {
            foreach (DataRow row in dt.Rows) 
            {
                if (row.RowState == DataRowState.Deleted) continue;
                string attachStr = row["附件檔案"]?.ToString();
                if (string.IsNullOrEmpty(attachStr)) continue;

                string rowDateStr = row[_dateColumnName]?.ToString() ?? "";
                string targetFolder = GetExpectedFolderName(rowDateStr); 

                string[] paths = attachStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                bool changed = false;
                
                for (int i = 0; i < paths.Length; i++) 
                {
                    string oldRelPath = paths[i].Replace("\\", "/");
                    string fileName = Path.GetFileName(oldRelPath);
                    string oldDir = Path.GetDirectoryName(oldRelPath).Replace("\\", "/");

                    string expectedRelDir = $"附件/{_dbName}/{_tableName}/{targetFolder}";

                    if (!oldDir.Equals(expectedRelDir, StringComparison.OrdinalIgnoreCase)) 
                    {
                        bool usedByOthersInGrid = false;
                        foreach(DataRow r in dt.Rows) {
                            if (r == row || r.RowState == DataRowState.Deleted) continue;
                            string otherAttach = r["附件檔案"]?.ToString();
                            if (!string.IsNullOrEmpty(otherAttach) && otherAttach.Contains(oldRelPath)) {
                                usedByOthersInGrid = true;
                                break;
                            }
                        }

                        bool usedByOthersInDb = false;
                        int currentRowId = -1;
                        if (dt.Columns.Contains("Id") && row["Id"] != DBNull.Value) {
                            int.TryParse(row["Id"].ToString(), out currentRowId);
                        }

                        try {
                            DataTable dbDt = DataManager.GetTableData(_dbName, _tableName, "", "", "");
                            foreach (DataRow dbRow in dbDt.Rows) {
                                int dbId = Convert.ToInt32(dbRow["Id"]);
                                if (dbId == currentRowId) continue;
                                string dbAttach = dbRow["附件檔案"]?.ToString();
                                if (!string.IsNullOrEmpty(dbAttach) && dbAttach.Contains(oldRelPath)) {
                                    usedByOthersInDb = true;
                                    break;
                                }
                            }
                        } catch { }

                        if (usedByOthersInGrid || usedByOthersInDb) {
                            continue; 
                        }

                        string oldAbsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, oldRelPath);
                        string newAbsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, expectedRelDir);
                        if (!Directory.Exists(newAbsDir)) Directory.CreateDirectory(newAbsDir);

                        string newAbsPath = Path.Combine(newAbsDir, fileName);
                        
                        int counter = 1;
                        string baseName = Path.GetFileNameWithoutExtension(fileName);
                        string ext = Path.GetExtension(fileName);
                        while (File.Exists(newAbsPath) && oldAbsPath != newAbsPath) 
                        {
                            fileName = $"{baseName}_{counter++}{ext}";
                            newAbsPath = Path.Combine(newAbsDir, fileName);
                        }

                        if (File.Exists(oldAbsPath)) 
                        {
                            File.Move(oldAbsPath, newAbsPath);
                            paths[i] = $"{expectedRelDir}/{fileName}";
                            changed = true;
                        }
                    }
                }
                
                if (changed) 
                {
                    row["附件檔案"] = string.Join("|", paths);
                }
            }
        }

        private async void BtnSave_Click(object sender, EventArgs e) 
        {
            try 
            {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _dgv.EndEdit(); 
                SaveColumnOrder(); 
                SetUIState(false, "資料庫寫入與檔案同步中，請稍候...", Color.Orange);
                
                DataTable dt = (DataTable)_dgv.DataSource;
                bool success = false;
                
                await Task.Run(() => { 
                    EnforceDateFormats(dt); 
                    SyncAttachmentPaths(dt);
                    success = DataManager.BulkSaveTable(_dbName, _tableName, dt); 
                });
                
                if (success) 
                { 
                    SetUIState(true, "資料儲存成功！", Color.Green); 
                    MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    await LoadGridDataAsync(); 
                } 
                else 
                { 
                    SetUIState(true, "資料儲存失敗", Color.Red); 
                }
            } 
            catch (Exception ex) 
            { 
                SetUIState(true, "儲存異常", Color.Red); 
                MessageBox.Show("儲存異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
            finally 
            { 
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
            }
        }

        private void BtnExport_Click(object sender, EventArgs e) 
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv", FileName = _chineseTitle + "_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (sfd.FilterIndex == 1) 
                        { 
                            using (ExcelPackage p = new ExcelPackage()) 
                            { 
                                var ws = p.Workbook.Worksheets.Add("Data"); 
                                ws.Cells["A1"].LoadFromDataTable(dt, true); 
                                p.SaveAs(new FileInfo(sfd.FileName)); 
                            } 
                        } 
                        else 
                        {
                            StringBuilder sb = new StringBuilder(); 
                            sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                            
                            foreach (DataRow r in dt.Rows) 
                            {
                                sb.AppendLine(string.Join(",", r.ItemArray.Select(i => i?.ToString().Replace(",", "，"))));
                            }
                            File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        }
                        MessageBox.Show("匯出成功！(附件欄位將輸出為相對路徑，以保證資料完整性)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
            }
        }

        private async void BtnImportExcel_Click(object sender, EventArgs e) 
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "請選擇要匯入的 Excel 檔案" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                        SetUIState(false, "Excel 解析與背景運算中，請稍候...", Color.Orange);

                        DataTable dt = (DataTable)_dgv.DataSource;
                        _dgv.DataSource = null; 
                        
                        await Task.Run(() => {
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) 
                            {
                                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                                if (ws == null || ws.Dimension == null) return;
                                
                                int rowCount = ws.Dimension.Rows; 
                                int colCount = ws.Dimension.Columns;
                                
                                string[] headers = new string[colCount];
                                for (int c = 1; c <= colCount; c++) 
                                {
                                    headers[c - 1] = ws.Cells[1, c].Text.Trim();
                                }

                                _calcHelper?.BeginBulkUpdate();
                                
                                for (int r = 2; r <= rowCount; r++) 
                                {
                                    DataRow nr = dt.NewRow(); 
                                    bool hasData = false;
                                    
                                    for (int c = 1; c <= colCount; c++) 
                                    {
                                        string cn = headers[c - 1]; 
                                        string val = ws.Cells[r, c].Text.Trim(); 
                                        
                                        if (dt.Columns.Contains(cn) && cn != "Id" && !string.IsNullOrEmpty(val)) 
                                        {
                                            nr[cn] = val; 
                                            hasData = true;
                                        }
                                    }
                                    if (hasData) dt.Rows.Add(nr);
                                }
                                
                                _calcHelper?.RecalculateTable(dt); 
                                _calcHelper?.EndBulkUpdate(); 
                                EnforceDateFormats(dt);
                            }
                        });
                        
                        _dgv.DataSource = dt; 
                        ApplyGridStyles(); 
                        RestoreColumnOrder();
                        SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{dt.Rows.Count}", Color.Green);
                        MessageBox.Show("Excel 匯入成功！\n請檢查數據後點擊「儲存數據」。", "匯入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    { 
                        await LoadGridDataAsync(); 
                        MessageBox.Show("匯入異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    } 
                    finally 
                    { 
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
                    }
                }
            }
        }

        private void Dgv_KeyDown(object sender, KeyEventArgs e) 
        {
            if (e.Control && e.KeyCode == Keys.V) 
            {
                try 
                {
                    string text = Clipboard.GetText(); 
                    if (string.IsNullOrEmpty(text)) return;
                    
                    _calcHelper?.BeginBulkUpdate();
                    string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    
                    int r = _dgv.CurrentCell.RowIndex;
                    int c = _dgv.CurrentCell.ColumnIndex;
                    DataTable dt = (DataTable)_dgv.DataSource;
                    
                    foreach (string line in lines) 
                    {
                        if (r >= _dgv.Rows.Count - 1) dt.Rows.Add(dt.NewRow());
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.Length; i++) 
                        {
                            if (c + i < _dgv.Columns.Count) 
                            {
                                if (_dgv.Columns[c + i].Name.Contains("附件檔案") || !_dgv.Columns[c + i].ReadOnly) 
                                {
                                    _dgv[c + i, r].Value = cells[i].Trim().Trim('"');
                                }
                            }
                        }
                        r++;
                    }
                    _calcHelper?.RecalculateTable(dt); 
                    _calcHelper?.EndBulkUpdate(); 
                    EnforceDateFormats(dt); 
                    _dgv.Refresh();
                } 
                catch 
                { 
                    _calcHelper?.EndBulkUpdate(); 
                }
            }
        }

        private class AttachmentForm : Form
        {
            public string FinalPathsString { get; private set; }
            private List<string> _paths = new List<string>();
            private string _dbName, _tableName, _targetFolder;
            private Action<string> _deleteAction;
            private FlowLayoutPanel _flpList;

            public AttachmentForm(string currentRelPathStr, string dbName, string tableName, string targetFolder, Action<string> deleteAction) 
            {
                _dbName = dbName; 
                _tableName = tableName; 
                _targetFolder = targetFolder; 
                _deleteAction = deleteAction;
                
                if (!string.IsNullOrEmpty(currentRelPathStr)) 
                {
                    _paths = new List<string>(currentRelPathStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries));
                }
                
                this.Text = "多檔附件管理中心"; 
                this.Size = new Size(700, 600); 
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.FixedDialog; 
                this.MaximizeBox = false; 
                this.MinimizeBox = false; 
                this.BackColor = Color.White;

                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F)); 
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); 
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));

                GroupBox boxList = new GroupBox { Text = "1. 已上傳檔案清單", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
                _flpList = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                boxList.Controls.Add(_flpList); 
                tlp.Controls.Add(boxList, 0, 0);

                GroupBox boxUpload = new GroupBox { Text = "2. 新增附件檔案", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(10) };
                Panel pnlDrop = new Panel { Dock = DockStyle.Fill, AllowDrop = true, BackColor = Color.AliceBlue, Cursor = Cursors.Hand };
                pnlDrop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlDrop.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Dashed);
                
                Label lblDrop = new Label { Text = "📁 點擊此處選擇多個檔案\n\n或\n\n將檔案拖曳至此區域", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), ForeColor = Color.SteelBlue };
                lblDrop.Click += (s, e) => SelectFiles(); 
                pnlDrop.Click += (s, e) => SelectFiles(); 
                pnlDrop.Controls.Add(lblDrop);
                
                pnlDrop.DragEnter += (s, e) => { 
                    if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; 
                };
                pnlDrop.DragDrop += (s, e) => { 
                    ProcessUpload((string[])e.Data.GetData(DataFormats.FileDrop)); 
                };
                boxUpload.Controls.Add(pnlDrop); 
                tlp.Controls.Add(boxUpload, 0, 1);

                Button btnClearAll = new Button { Text = "🗑️ 清除此筆紀錄的所有附件", Dock = DockStyle.Fill, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(3, 5, 3, 5) };
                btnClearAll.Click += (s, e) => {
                    if (_paths.Count == 0) return;
                    if (MessageBox.Show("確定要清除所有附件嗎？\n(實體檔案將被同步永久刪除)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                    {
                        foreach (var p in _paths) _deleteAction(p);
                        _paths.Clear(); 
                        RefreshListUI();
                    }
                };
                tlp.Controls.Add(btnClearAll, 0, 2);

                Button btnSaveClose = new Button { Text = "💾 確認變更並返回", Dock = DockStyle.Fill, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), Margin = new Padding(3, 5, 3, 5) };
                btnSaveClose.Click += (s, e) => { 
                    FinalPathsString = string.Join("|", _paths); 
                    this.DialogResult = DialogResult.OK; 
                };
                tlp.Controls.Add(btnSaveClose, 0, 3);

                this.Controls.Add(tlp); 
                RefreshListUI();
            }

            private void RefreshListUI() 
            {
                _flpList.Controls.Clear();
                if (_paths.Count == 0) 
                { 
                    _flpList.Controls.Add(new Label { Text = "(尚無任何附件)", ForeColor = Color.DimGray, AutoSize = true, Margin = new Padding(10) }); 
                    return; 
                }
                
                foreach (string path in _paths) 
                {
                    Panel pItem = new Panel { Width = _flpList.Width - 30, Height = 40, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                    Label lName = new Label { Text = Path.GetFileName(path), Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Microsoft JhengHei UI", 11F) };
                    
                    Button bOpen = new Button { Text = "開啟", Width = 80, Dock = DockStyle.Right, BackColor = Color.LightGray, Cursor = Cursors.Hand };
                    bOpen.Click += (s, e) => { 
                        try { System.Diagnostics.Process.Start(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path)); } 
                        catch (Exception ex) { MessageBox.Show("開啟失敗：" + ex.Message); } 
                    };

                    Button bDownload = new Button { Text = "下載", Width = 80, Dock = DockStyle.Right, BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
                    bDownload.Click += (s, e) => {
                        try 
                        {
                            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
                            if (!File.Exists(sourcePath)) 
                            {
                                MessageBox.Show("找不到原始檔案，可能已被移動或刪除。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            using (SaveFileDialog sfd = new SaveFileDialog())
                            {
                                string fileName = Path.GetFileName(path);
                                string ext = Path.GetExtension(path);
                                sfd.FileName = fileName;
                                sfd.Title = "另存附件檔案";
                                sfd.Filter = $"檔案 (*{ext})|*{ext}|所有檔案 (*.*)|*.*";
                                
                                if (sfd.ShowDialog() == DialogResult.OK)
                                {
                                    File.Copy(sourcePath, sfd.FileName, true);
                                    MessageBox.Show("檔案下載完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        catch (Exception ex) 
                        {
                            MessageBox.Show("下載失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };
                    
                    Button bDel = new Button { Text = "刪除", Width = 80, Dock = DockStyle.Right, BackColor = Color.LightCoral, ForeColor = Color.White, Cursor = Cursors.Hand };
                    bDel.Click += (s, e) => { 
                        if (MessageBox.Show($"確定刪除 {Path.GetFileName(path)}?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) 
                        { 
                            _deleteAction(path); 
                            _paths.Remove(path); 
                            RefreshListUI(); 
                        } 
                    };
                    
                    pItem.Controls.Add(lName); 
                    pItem.Controls.Add(bDel);       
                    pItem.Controls.Add(bDownload);  
                    pItem.Controls.Add(bOpen);      
                    
                    _flpList.Controls.Add(pItem);
                }
            }

            private void SelectFiles() 
            {
                using (OpenFileDialog ofd = new OpenFileDialog { Title = "選擇附件檔案", Multiselect = true, Filter = "所有檔案 (*.*)|*.*" }) 
                {
                    if (ofd.ShowDialog() == DialogResult.OK) ProcessUpload(ofd.FileNames);
                }
            }

            private void ProcessUpload(string[] sourceFiles) 
            {
                if (sourceFiles.Length == 0) return;
                
                string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件", _dbName, _tableName, _targetFolder);
                
                if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
                
                foreach (string src in sourceFiles) 
                {
                    try 
                    {
                        string ext = Path.GetExtension(src); 
                        string baseName = Path.GetFileNameWithoutExtension(src);
                        string destName = baseName + ext; 
                        string destPath = Path.Combine(destDir, destName);
                        
                        int count = 1; 
                        while (File.Exists(destPath)) 
                        { 
                            destName = $"{baseName}_{count++}{ext}"; 
                            destPath = Path.Combine(destDir, destName); 
                        }
                        
                        File.Copy(src, destPath); 
                        _paths.Add($"附件/{_dbName}/{_tableName}/{_targetFolder}/{destName}");
                    } 
                    catch (Exception ex) 
                    { 
                        MessageBox.Show($"上傳檔案 {Path.GetFileName(src)} 失敗: {ex.Message}", "錯誤"); 
                    }
                }
                RefreshListUI();
            }
        }
    }
}
