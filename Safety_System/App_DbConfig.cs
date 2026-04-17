using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_DbConfig
    {
        private TextBox _txtPath;
        private TextBox _txtBackupPath;
        private NumericUpDown _numKeepCount;
        private ComboBox _cboDb, _cboTable, _cboCol1, _cboCol2, _cboCol3, _cboCol4;

        // 🟢 用於中文化顯示的字典結構
        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => ChName; // UI 顯示中文
        }

        // 定義資料庫與資料表的中文對照關係
        private readonly Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap = new Dictionary<string, (string, Dictionary<string, string>)> {
            { "Safety", ("工安", new Dictionary<string, string> { 
                { "NearMiss", "虛驚事件" }, { "SafetyInspection", "巡檢記錄" }, { "SafetyObservation", "安全觀察" }, 
                { "TrafficInjury", "交通意外" }, { "WorkInjury", "工傷事件" }, { "MinorInjury", "輕傷事件" } 
            })},
            { "Chemical", ("化學品", new Dictionary<string, string> { 
                { "SDS_Inventory", "SDS清冊" }, { "EnvTesting", "環測項目" }, { "ExposureLimits", "勞工暴露容許濃度" }, 
                { "ToxicSubstances", "毒性物質" }, { "ConcernedChem", "關注性化學物質" }, { "PriorityMgmtChem", "優先管理化學品" }, 
                { "ControlledChem", "管制化學品" }, { "SpecificChem", "特定化學物質" }, { "OrganicSolvents", "有機溶劑" }, 
                { "WorkerHealthProtect", "勞工健康保護" }, { "PublicHazardous", "公共危險物品" }, { "AirPollutionEmerg", "空污緊急應變" }, 
                { "FactoryHazardous", "工廠危險物品申報" } 
            })},
            { "Nursing", ("護理", new Dictionary<string, string> { { "HealthPromotion", "健康促進活動" }, { "WorkInjuryReport", "職災申報紀錄" } })},
            { "Air", ("空污", new Dictionary<string, string> { { "AirPollution", "空污申報紀錄" } })},
            { "Water", ("水污", new Dictionary<string, string> { 
                { "DischargeData", "納管排放數據" }, { "WaterMeterReadings", "廢水處理水量記錄" }, { "WaterChemicals", "廢水處理用藥記錄" }, 
                { "WaterVolume", "自來水用量統計" }, { "WaterUsageDaily", "自來水使用量" } 
            })},
            { "Waste", ("產能及廢棄物", new Dictionary<string, string> { 
                { "WasteMonthly", "廢棄物月表" }, { "Waste_IL", "複層月表" }, { "Waste_LM", "膠合月表" }, { "Waste_CR", "鍍板月表" }, 
                { "Waste_T", "強化月表" }, { "Waste_GCTE", "切磨月表" }, { "Waste_ML", "物料月表" }, { "Waste_Water", "水站月表" } 
            })},
            { "Fire", ("消防", new Dictionary<string, string> { 
                { "FireResponsible", "火源責任人" }, { "HazardStats", "公共危險物統計" }, { "FireEquip", "消防設備巡檢" }, { "FireSelfInspection", "各單位消防自主檢查" } 
            })},
            { "TestData", ("檢測數據", new Dictionary<string, string> { 
                { "CoolingWaterSelf", "循環水檢測(自評)" }, { "CoolingWaterVendor", "循環水檢測(廠商)" }, { "DrinkingWater", "飲用水檢測" }, 
                { "EnvMonitor", "環境監測" }, { "IndustrialZoneTest", "工業區檢驗" }, { "OtherTests", "其它檢測數據" }, 
                { "SoilGasTest", "土壤氣體檢測" }, { "TCLP", "TCLP毒性特性溶出" }, { "WastewaterPeriodic", "廢水定申檢" }, 
                { "WastewaterSelfTest", "廢水自主檢驗" }, { "WaterMeterCalibration", "水錶校正" } 
            })},
            { "教育訓練", ("教育訓練", new Dictionary<string, string> { { "訓練時數", "教育訓練時數" } })},
            { "法規", ("法規", new Dictionary<string, string> { { "環保法規", "環保法規" }, { "職安衛法規", "職安衛法規" }, { "消防法規", "消防法規" }, { "其它法規", "其它法規" } })}, 
            { "ESG", ("ESG", new Dictionary<string, string> { { "ESG_Performance", "ESG績效管理" } })},
            { "ISO14001", ("ISO14001", new Dictionary<string, string> { { "TargetManagement", "目標管理" } })},
            { "Purchase", ("請購", new Dictionary<string, string> { { "PurchaseData", "請購資料" } })} 
        };

        public Control GetView()
        {
            Panel main = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

            // ==========================================
            // 1. 資料庫存放路徑設定
            // ==========================================
            GroupBox boxPath = new GroupBox { Text = "資料庫存放路徑設定", Dock = DockStyle.Top, Height = 180, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            string currentPath = string.IsNullOrEmpty(DataManager.BasePath) ? "" : DataManager.BasePath;
            _txtPath = new TextBox { Location = new Point(30, 50), Width = 600, ReadOnly = true, Text = currentPath, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Button btnBrowse = new Button { Text = "選擇資料夾", Location = new Point(650, 48), Size = new Size(150, 35), Font = new Font("Microsoft JhengHei UI", 12F) };
            btnBrowse.Click += (s, e) => {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "請選擇數據資料存放的資料夾" }) {
                    if (fbd.ShowDialog() == DialogResult.OK) _txtPath.Text = fbd.SelectedPath;
                }
            };
            
            Button btnSavePath = new Button { Text = "儲存路徑變更", Location = new Point(30, 110), Size = new Size(220, 45), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSavePath.Click += (s, e) => {
                if (!AuthManager.VerifyAdmin()) return; 
                if (System.IO.Directory.Exists(_txtPath.Text)) {
                    DataManager.SetBasePath(_txtPath.Text);
                    MessageBox.Show("路徑已更新！後續系統存取皆會依此路徑。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("請選擇有效的資料夾路徑。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            boxPath.Controls.AddRange(new Control[] { _txtPath, btnBrowse, btnSavePath });

            // ==========================================
            // 2. 資料備份設定
            // ==========================================
            GroupBox boxBackup = new GroupBox { Text = "資料庫備份設定 (自動每週備份)", Dock = DockStyle.Top, Height = 220, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            boxBackup.Margin = new Padding(0, 30, 0, 0);

            BackupManager.LoadConfig();

            Label lblB1 = new Label { Text = "備份存放路徑:", Location = new Point(30, 50), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _txtBackupPath = new TextBox { Location = new Point(200, 47), Width = 430, ReadOnly = true, Text = BackupManager.BackupPath, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Button btnBrowseBackup = new Button { Text = "選擇資料夾", Location = new Point(650, 45), Size = new Size(150, 35), Font = new Font("Microsoft JhengHei UI", 12F) };
            btnBrowseBackup.Click += (s, e) => {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "請選擇備份資料存放的資料夾" }) {
                    if (fbd.ShowDialog() == DialogResult.OK) _txtBackupPath.Text = fbd.SelectedPath;
                }
            };

            Label lblB2 = new Label { Text = "保留舊備份份數:", Location = new Point(30, 100), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _numKeepCount = new NumericUpDown { Location = new Point(200, 98), Width = 80, Minimum = 1, Maximum = 100, Value = BackupManager.KeepCount, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblB3 = new Label { Text = "份 (建議保留 4 份，約一個月)", Location = new Point(290, 100), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };

            Button btnSaveBackup = new Button { Text = "儲存備份設定", Location = new Point(30, 150), Size = new Size(220, 45), BackColor = Color.Sienna, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveBackup.Click += (s, e) => {
                if (!AuthManager.VerifyAdmin()) return;
                BackupManager.SaveConfig(_txtBackupPath.Text, (int)_numKeepCount.Value);
                MessageBox.Show("備份設定已儲存！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Button btnManualBackup = new Button { Text = "立即執行手動備份", Location = new Point(300, 150), Size = new Size(200, 45), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnManualBackup.Click += (s, e) => {
                BackupManager.ExecuteBackup();
                MessageBox.Show("手動備份執行完成！", "備份成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            boxBackup.Controls.AddRange(new Control[] { lblB1, _txtBackupPath, btnBrowseBackup, lblB2, _numKeepCount, lblB3, btnSaveBackup, btnManualBackup });

            // ==========================================
            // 3. 資料表防重寫欄位設定 (🟢 高度拉大至 400，容納 4 個欄位)
            // ==========================================
            GroupBox boxKeys = new GroupBox { Text = "資料表防重寫欄位設定 (空值則正常寫入不防呆)", Dock = DockStyle.Top, Height = 400, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            boxKeys.Margin = new Padding(0, 30, 0, 0);

            Label lblDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboDb = new ComboBox { Location = new Point(160, 58), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblTable = new Label { Text = "選擇資料表:", Location = new Point(420, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable = new ComboBox { Location = new Point(540, 58), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol1 = new Label { Text = "判斷欄位一:", Location = new Point(30, 130), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol1 = new ComboBox { Location = new Point(160, 128), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol2 = new Label { Text = "判斷欄位二:", Location = new Point(420, 130), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol2 = new ComboBox { Location = new Point(540, 128), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            // 🟢 新增第三與第四個判斷欄位
            Label lblCol3 = new Label { Text = "判斷欄位三:", Location = new Point(30, 200), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol3 = new ComboBox { Location = new Point(160, 198), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol4 = new Label { Text = "判斷欄位四:", Location = new Point(420, 200), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol4 = new ComboBox { Location = new Point(540, 198), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnSaveKeys = new Button { Text = "儲存防重寫規則", Location = new Point(30, 280), Size = new Size(220, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveKeys.Click += BtnSaveKeys_Click;

            boxKeys.Controls.AddRange(new Control[] { 
                lblDb, _cboDb, lblTable, _cboTable, 
                lblCol1, _cboCol1, lblCol2, _cboCol2, 
                lblCol3, _cboCol3, lblCol4, _cboCol4, 
                btnSaveKeys 
            });

            Panel spacer1 = new Panel { Dock = DockStyle.Top, Height = 30 };
            Panel spacer2 = new Panel { Dock = DockStyle.Top, Height = 30 };

            main.Controls.Add(boxKeys);
            main.Controls.Add(spacer1);
            main.Controls.Add(boxBackup);
            main.Controls.Add(spacer2);
            main.Controls.Add(boxPath);

            // 載入中文化資料庫選項
            foreach (var kvp in _dbMap) {
                _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
            }
            
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
            
            if (_cboDb.Items.Count > 0) _cboDb.SelectedIndex = 0;

            return main;
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboTable == null || _cboCol1 == null || _cboCol2 == null || _cboCol3 == null || _cboCol4 == null) return;
                _cboTable.Items.Clear(); _cboCol1.Items.Clear(); _cboCol2.Items.Clear(); _cboCol3.Items.Clear(); _cboCol4.Items.Clear();
                if (_cboDb.SelectedItem == null) return;
                
                var selectedDb = (ItemMap)_cboDb.SelectedItem;
                if (_dbMap.ContainsKey(selectedDb.EnName)) {
                    foreach (var tbl in _dbMap[selectedDb.EnName].Tables) {
                        _cboTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                    }
                    if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
                }
            } catch { }
        }

        private void CboTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboCol1 == null) return;
                _cboCol1.Items.Clear(); _cboCol2.Items.Clear(); _cboCol3.Items.Clear(); _cboCol4.Items.Clear();
                _cboCol1.Items.Add(""); _cboCol2.Items.Add(""); _cboCol3.Items.Add(""); _cboCol4.Items.Add("");
                
                if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) return;
                
                string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
                string tableName = ((ItemMap)_cboTable.SelectedItem).EnName;

                List<string> cols = DataManager.GetColumnNames(dbName, tableName);
                foreach(var c in cols) if (c != "Id") { 
                    _cboCol1.Items.Add(c); _cboCol2.Items.Add(c); _cboCol3.Items.Add(c); _cboCol4.Items.Add(c);
                }

                var keys = DataManager.GetTableKeys(dbName, tableName);
                
                if (!string.IsNullOrEmpty(keys.col1) && _cboCol1.Items.Contains(keys.col1)) _cboCol1.SelectedItem = keys.col1; else _cboCol1.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(keys.col2) && _cboCol2.Items.Contains(keys.col2)) _cboCol2.SelectedItem = keys.col2; else _cboCol2.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(keys.col3) && _cboCol3.Items.Contains(keys.col3)) _cboCol3.SelectedItem = keys.col3; else _cboCol3.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(keys.col4) && _cboCol4.Items.Contains(keys.col4)) _cboCol4.SelectedItem = keys.col4; else _cboCol4.SelectedIndex = 0;
            } catch { }
        }

        private void BtnSaveKeys_Click(object sender, EventArgs e)
        {
            if (!AuthManager.VerifyAdmin()) return; 

            if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) {
                MessageBox.Show("請先選擇資料庫與資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            
            string dbName = ((ItemMap)_cboDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboTable.SelectedItem).EnName;
            string chTableName = ((ItemMap)_cboTable.SelectedItem).ChName;
            
            string c1 = _cboCol1.SelectedItem?.ToString() ?? "";
            string c2 = _cboCol2.SelectedItem?.ToString() ?? "";
            string c3 = _cboCol3.SelectedItem?.ToString() ?? "";
            string c4 = _cboCol4.SelectedItem?.ToString() ?? "";

            DataManager.SaveTableKeys(dbName, tableName, c1, c2, c3, c4);
            MessageBox.Show($"【{chTableName}】 防重寫規則儲存成功！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
