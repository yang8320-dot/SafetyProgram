/// FILE: Safety_System/App_DbConfig.cs ///
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
        private ComboBox _cboDb, _cboTable, _cboCol1, _cboCol2;

        // 🟢 加入 MinorInjury 到 Safety，加入 ISO14001
        private readonly Dictionary<string, string[]> _dbMap = new Dictionary<string, string[]> {
            { "Safety", new[] { "NearMiss", "SafetyInspection", "SafetyObservation", "TrafficInjury", "WorkInjury", "MinorInjury" } },
            { "Chemical", new[] { "ChemRegulations", "SDS_Inventory" } },
            { "Nursing", new[] { "HealthPromotion", "WorkInjuryReport" } },
            { "Air", new[] { "AirPollution" } },
            { "Water", new[] { "DischargeData", "WaterMeterReadings", "WaterChemicals", "WaterVolume", "WaterUsageDaily" } },
            { "Waste", new[] { "WasteMonthly", "Waste_IL", "Waste_LM", "Waste_CR", "Waste_T", "Waste_GCTE", "Waste_ML", "Waste_Water" } },
            { "Fire", new[] { "FireResponsible", "HazardStats", "FireEquip" } },
            { "TestData", new[] { "CoolingWaterSelf", "CoolingWaterVendor", "DrinkingWater", "EnvMonitor", "IndustrialZoneTest", "OtherTests", "SoilGasTest", "TCLP", "WastewaterPeriodic", "WastewaterSelfTest", "WaterMeterCalibration" } },
            { "教育訓練", new[] { "訓練時數" } },
            { "法規", new[] { "環保法規", "職安衛法規", "其它法規" } },
            { "ESG", new[] { "ESG_Performance" } },
            { "ISO14001", new[] { "TargetManagement" } } 
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
            // 3. 資料表防重寫欄位設定
            // ==========================================
            GroupBox boxKeys = new GroupBox { Text = "資料表防重寫欄位設定 (空值則正常寫入不防呆)", Dock = DockStyle.Top, Height = 320, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            boxKeys.Margin = new Padding(0, 30, 0, 0);

            Label lblDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboDb = new ComboBox { Location = new Point(160, 58), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblTable = new Label { Text = "選擇資料表:", Location = new Point(420, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable = new ComboBox { Location = new Point(540, 58), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol1 = new Label { Text = "判斷欄位一:", Location = new Point(30, 130), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol1 = new ComboBox { Location = new Point(160, 128), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol2 = new Label { Text = "判斷欄位二:", Location = new Point(420, 130), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol2 = new ComboBox { Location = new Point(540, 128), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnSaveKeys = new Button { Text = "儲存防重寫規則", Location = new Point(30, 210), Size = new Size(220, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveKeys.Click += BtnSaveKeys_Click;

            boxKeys.Controls.AddRange(new Control[] { lblDb, _cboDb, lblTable, _cboTable, lblCol1, _cboCol1, lblCol2, _cboCol2, btnSaveKeys });

            Panel spacer1 = new Panel { Dock = DockStyle.Top, Height = 30 };
            Panel spacer2 = new Panel { Dock = DockStyle.Top, Height = 30 };

            // 依序反著加進去 (Dock.Top 特性)
            main.Controls.Add(boxKeys);
            main.Controls.Add(spacer1);
            main.Controls.Add(boxBackup);
            main.Controls.Add(spacer2);
            main.Controls.Add(boxPath);

            foreach (var key in _dbMap.Keys) _cboDb.Items.Add(key);
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
            
            if (_cboDb.Items.Count > 0) _cboDb.SelectedIndex = 0;

            return main;
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboTable == null || _cboCol1 == null || _cboCol2 == null) return;
                _cboTable.Items.Clear(); _cboCol1.Items.Clear(); _cboCol2.Items.Clear();
                if (_cboDb.SelectedItem == null) return;
                
                string db = _cboDb.SelectedItem.ToString();
                if (_dbMap.ContainsKey(db)) {
                    _cboTable.Items.AddRange(_dbMap[db]);
                    if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
                }
            } catch { }
        }

        private void CboTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboCol1 == null || _cboCol2 == null) return;
                _cboCol1.Items.Clear(); _cboCol2.Items.Clear();
                _cboCol1.Items.Add(""); _cboCol2.Items.Add(""); 
                if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) return;
                
                string dbName = _cboDb.SelectedItem.ToString();
                string tableName = _cboTable.SelectedItem.ToString();

                List<string> cols = DataManager.GetColumnNames(dbName, tableName);
                foreach(var c in cols) if (c != "Id") { _cboCol1.Items.Add(c); _cboCol2.Items.Add(c); }

                var keys = DataManager.GetTableKeys(dbName, tableName);
                if (!string.IsNullOrEmpty(keys.col1) && _cboCol1.Items.Contains(keys.col1)) _cboCol1.SelectedItem = keys.col1; else _cboCol1.SelectedIndex = 0;
                if (!string.IsNullOrEmpty(keys.col2) && _cboCol2.Items.Contains(keys.col2)) _cboCol2.SelectedItem = keys.col2; else _cboCol2.SelectedIndex = 0;
            } catch { }
        }

        private void BtnSaveKeys_Click(object sender, EventArgs e)
        {
            if (!AuthManager.VerifyAdmin()) return; 

            if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) {
                MessageBox.Show("請先選擇資料庫與資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            string dbName = _cboDb.SelectedItem.ToString();
            string tableName = _cboTable.SelectedItem.ToString();
            string c1 = _cboCol1.SelectedItem?.ToString() ?? "";
            string c2 = _cboCol2.SelectedItem?.ToString() ?? "";

            DataManager.SaveTableKeys(dbName, tableName, c1, c2);
            MessageBox.Show($"【{tableName}】 防重寫規則儲存成功！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
