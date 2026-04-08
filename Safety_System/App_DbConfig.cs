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
        private ComboBox _cboDb, _cboTable, _cboCol1, _cboCol2;

        // 🟢 完整涵蓋原有模組，並加上教育訓練與 61 個法規資料表
        private readonly Dictionary<string, string[]> _dbMap = new Dictionary<string, string[]> {
            { "Safety", new[] { "NearMiss", "SafetyInspection", "SafetyObservation", "TrafficInjury", "WorkInjury" } },
            { "Nursing", new[] { "HealthPromotion", "WorkInjuryReport" } },
            { "Air", new[] { "AirPollution" } },
            { "Water", new[] { "DischargeData", "WaterMeterReadings", "WaterChemicals", "WaterVolume", "WaterUsageDaily" } },
            { "Waste", new[] { "WasteMonthly" } },
            { "Fire", new[] { "FireResponsible", "HazardStats", "FireEquip" } },
            { "TestData", new[] { "CoolingWaterSelf", "CoolingWaterVendor", "DrinkingWater", "EnvMonitor", "IndustrialZoneTest", "OtherTests", "SoilGasTest", "TCLP", "WastewaterPeriodic", "WastewaterSelfTest", "WaterMeterCalibration" } },
            
            // =========== 新增 ===========
            { "教育訓練", new[] { "訓練時數" } },
            { "環保法規", new[] { "組織與處務", "環境綜合計畫", "環境影響評估", "空氣污染防制", "噪音污染管制", "水污染防治", "海洋污染防制", "廢棄物清理", "應回收廢棄物", "資源回收再利用", "土壤污染整治", "毒化物管理", "飲用水管理", "環境用藥管理", "公害糾紛處理", "環境污染檢驗", "環保人員訓練", "環境教育", "溫室氣體管理", "室內空氣管理" } },
            { "職安衛法規", new[] { "一般安全衛生法規", "一般環境管理相關法規", "高壓氣體相關法規", "健康管理相關法規", "教育訓練相關法規", "化學物質相關法規", "機械安全相關法規", "特殊作業相關法規", "特別行業適用法規", "職業災害勞工保護法相關法規", "安衛其他法規", "勞資關係", "勞動條件", "勞工福利", "勞工保險", "職業訓練", "就業服務" } },
            { "其它法規", new[] { "原子能相關法規", "衛生福利相關法規", "交通安全相關法規", "消防相關法規", "建築相關法規", "下水道相關法規", "科學園區相關法規", "一般工業區相關法規", "利害相關者要求", "國際環保公約", "水利相關法規", "能源管理", "文化相關法規", "族群相關法規", "消費相關法規", "財政金融相關法規", "法務相關法規", "社福警政相關法規", "經濟相關法規", "觀光旅遊相關法規", "生態相關法規", "礦業相關法規", "農林漁牧相關法規", "教育體育相關法規", "通訊傳播相關法規" } }
        };

        public Control GetView()
        {
            Panel main = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

            // 第一區塊：資料庫存放路徑設定
            GroupBox boxPath = new GroupBox { Text = "資料庫存放路徑設定", Dock = DockStyle.Top, Height = 180, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            string currentPath = string.IsNullOrEmpty(DataManager.BasePath) ? "" : DataManager.BasePath;
            _txtPath = new TextBox { Location = new Point(30, 50), Width = 600, ReadOnly = true, Text = currentPath, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Button btnBrowse = new Button { Text = "選擇資料夾", Location = new Point(650, 48), Size = new Size(150, 35), Font = new Font("Microsoft JhengHei UI", 12F) };
            btnBrowse.Click += BtnBrowse_Click;
            
            Button btnSavePath = new Button { Text = "儲存路徑變更", Location = new Point(30, 110), Size = new Size(220, 45), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSavePath.Click += BtnSavePath_Click;

            boxPath.Controls.AddRange(new Control[] { _txtPath, btnBrowse, btnSavePath });

            // 第二區塊：防重寫欄位設定
            GroupBox boxKeys = new GroupBox { Text = "資料表防重寫欄位設定 (空值則正常寫入不防呆)", Dock = DockStyle.Top, Height = 320, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            boxKeys.Margin = new Padding(0, 30, 0, 0);

            Label lblDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboDb = new ComboBox { Location = new Point(150, 58), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblTable = new Label { Text = "選擇資料表:", Location = new Point(420, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable = new ComboBox { Location = new Point(540, 58), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol1 = new Label { Text = "判斷欄位一:", Location = new Point(30, 130), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol1 = new ComboBox { Location = new Point(150, 128), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol2 = new Label { Text = "判斷欄位二:", Location = new Point(420, 130), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol2 = new ComboBox { Location = new Point(540, 128), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnSaveKeys = new Button { Text = "儲存防重寫規則", Location = new Point(30, 210), Size = new Size(220, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveKeys.Click += BtnSaveKeys_Click;

            boxKeys.Controls.AddRange(new Control[] { lblDb, _cboDb, lblTable, _cboTable, lblCol1, _cboCol1, lblCol2, _cboCol2, btnSaveKeys });

            Panel spacer = new Panel { Dock = DockStyle.Top, Height = 30 };

            main.Controls.Add(boxKeys);
            main.Controls.Add(spacer);
            main.Controls.Add(boxPath);

            foreach (var key in _dbMap.Keys) _cboDb.Items.Add(key);
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
            
            if (_cboDb.Items.Count > 0) _cboDb.SelectedIndex = 0;

            return main;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "請選擇數據資料存放的資料夾" })
            {
                if (fbd.ShowDialog() == DialogResult.OK) _txtPath.Text = fbd.SelectedPath;
            }
        }

        private void BtnSavePath_Click(object sender, EventArgs e)
        {
            if (System.IO.Directory.Exists(_txtPath.Text)) {
                DataManager.SetBasePath(_txtPath.Text);
                MessageBox.Show("路徑已更新！後續系統存取皆會依此路徑。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } else {
                MessageBox.Show("請選擇有效的資料夾路徑。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboTable == null || _cboCol1 == null || _cboCol2 == null) return;

                _cboTable.Items.Clear();
                _cboCol1.Items.Clear(); _cboCol2.Items.Clear();
                if (_cboDb.SelectedItem == null) return;
                
                string db = _cboDb.SelectedItem.ToString();
                if (_dbMap.ContainsKey(db)) {
                    _cboTable.Items.AddRange(_dbMap[db]);
                    if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
                }
            } catch { /* 攔截聯動錯誤 */ }
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
                foreach(var c in cols) {
                    if (c != "Id") { _cboCol1.Items.Add(c); _cboCol2.Items.Add(c); }
                }

                var keys = DataManager.GetTableKeys(dbName, tableName);
                if (!string.IsNullOrEmpty(keys.col1) && _cboCol1.Items.Contains(keys.col1)) _cboCol1.SelectedItem = keys.col1;
                else _cboCol1.SelectedIndex = 0;

                if (!string.IsNullOrEmpty(keys.col2) && _cboCol2.Items.Contains(keys.col2)) _cboCol2.SelectedItem = keys.col2;
                else _cboCol2.SelectedIndex = 0;
            } catch { /* 攔截聯動錯誤 */ }
        }

        private void BtnSaveKeys_Click(object sender, EventArgs e)
        {
            if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) {
                MessageBox.Show("請先選擇資料庫與資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            string dbName = _cboDb.SelectedItem.ToString();
            string tableName = _cboTable.SelectedItem.ToString();
            string c1 = _cboCol1.SelectedItem?.ToString() ?? "";
            string c2 = _cboCol2.SelectedItem?.ToString() ?? "";

            DataManager.SaveTableKeys(dbName, tableName, c1, c2);
            MessageBox.Show($"【{tableName}】 防重寫規則儲存成功！\n\n判定欄位1：{(c1==""?"無":c1)}\n判定欄位2：{(c2==""?"無":c2)}", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
