/// FILE: Safety_System/App_DbConfig.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_DbConfig
    {
        private TextBox _txtPath;
        private TextBox _txtBackupPath;
        private NumericUpDown _numKeepCount;
        
        private ComboBox _cboDb, _cboTable, _cboCol1, _cboCol2, _cboCol3, _cboCol4;
        private ComboBox _cboDelDb, _cboDelTable;

        private class SyncRowUI {
            public ComboBox CboSrcDb, CboSrcTable, CboSrcMatchCol, CboSrcSyncCol;
            public ComboBox CboTgtDb, CboTgtTable, CboTgtMatchCol, CboTgtSyncCol;
            public ComboBox CboSyncType; 
        }
        private List<SyncRowUI> _syncRows = new List<SyncRowUI>();

        private class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        private readonly Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap = new Dictionary<string, (string, Dictionary<string, string>)> {
            { "Safety", ("工安", new Dictionary<string, string> { 
                { "NearMiss", "虛驚事件" }, { "SafetyInspection", "巡檢記錄" }, { "SafetyObservation", "安全觀察" }, 
                { "TrafficInjury", "交通意外" }, { "WorkInjury", "工傷事件" }, { "MinorInjury", "輕傷事件" },
                { "LaborInspection", "勞檢稽查缺失" } 
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
                { "WaterVolume", "自來水用量統計" }, { "WaterUsageDaily", "自來水使用量" },
                { "WaterPermitMaterial", "水污許可(原物料)" } 
            })},
            { "Waste", ("廢棄物", new Dictionary<string, string> { 
                { "WasteMonthly", "廢棄物月表" }, { "Waste_IL", "複層月表" }, { "Waste_LM", "膠合月表" }, { "Waste_CR", "鍍板月表" }, 
                { "Waste_T", "強化月表" }, { "Waste_GCTE", "切磨月表" }, { "Waste_ML", "物料月表" }, { "Waste_Water", "水站月表" },
                { "WastePermitMaterial", "廢棄物污許可(原物料)" }, 
                { "WastePermitProduct", "廢棄物污許可(產品)" }, 
                { "WastePermitWaste", "廢棄物污許可(廢棄物)" } 
            })},
            { "Fire", ("消防", new Dictionary<string, string> { 
                { "FireResponsible", "火源責任人" }, { "HazardStats", "公共危險物統計" }, { "FireEquip", "消防設備清單" }, { "FireSelfInspection", "各單位消防自主檢查" } 
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
            { "ISO14001", ("ISO14001", new Dictionary<string, string> { 
                { "TargetManagement", "目標管理" }, 
                { "EnvInfoReceive", "環境資訊接收管制表" }, 
                { "InternalComm", "內文聯絡書管制表" }, 
                { "MailReceive", "郵件收文管制表" }, 
                { "VisitorRecord", "來賓拜訪紀錄表" }
            })},
            { "Purchase", ("請購", new Dictionary<string, string> { { "PurchaseData", "請購資料" } })},
            { "Menu1DB", ("選單1", new Dictionary<string, string> { { "WorkItems", "WorkItems" }, { "AccountManage", "帳密管理" }, { "KPI", "KPI" }, { "CultureImprove", "文化改善" }, { "PBC", "PBC" } })},
            { "Menu2DB", ("選單2", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })},
            { "Menu3DB", ("選單3", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })},
            { "Menu4DB", ("選單4", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })}
        };

        public Control GetView()
        {
            try {
                DataTable dtMenus = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
                if (dtMenus != null) {
                    foreach (DataRow r in dtMenus.Rows) {
                        string dbName = r["資料庫名"].ToString();
                        string tableName = r["資料表名"].ToString();
                        if (_dbMap.ContainsKey(dbName)) {
                            _dbMap[dbName].Tables[tableName] = tableName;
                        }
                    }
                }
            } catch { }

            Panel main = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(0,0,0,50) };

            // 1. 路徑設定區塊
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

            // 2. 備份區塊
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

            // 3. 防重寫規則區塊
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

            Label lblCol3 = new Label { Text = "判斷欄位三:", Location = new Point(30, 200), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol3 = new ComboBox { Location = new Point(160, 198), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol4 = new Label { Text = "判斷欄位四:", Location = new Point(420, 200), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol4 = new ComboBox { Location = new Point(540, 198), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnSaveKeys = new Button { Text = "儲存防重寫規則", Location = new Point(30, 280), Size = new Size(220, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveKeys.Click += BtnSaveKeys_Click;

            // 🟢 新增：防重寫設定清單 按鈕
            Button btnShowTableKeysList = new Button { Text = "防重寫設定清單", Location = new Point(280, 280), Size = new Size(200, 45), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnShowTableKeysList.Click += BtnShowTableKeysList_Click;

            boxKeys.Controls.AddRange(new Control[] { 
                lblDb, _cboDb, lblTable, _cboTable, 
                lblCol1, _cboCol1, lblCol2, _cboCol2, 
                lblCol3, _cboCol3, lblCol4, _cboCol4, 
                btnSaveKeys, btnShowTableKeysList 
            });

            // 4. 資料同步區塊
            GroupBox boxSync = new GroupBox { Text = "資料同步設定 (來源儲存時自動聚合計算至目標表)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            boxSync.Margin = new Padding(0, 30, 0, 0);

            TableLayoutPanel tlpSync = new TableLayoutPanel {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 10,
                Padding = new Padding(15, 20, 15, 10)
            };
            
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40F)); 
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11F));
            tlpSync.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 12F)); 

            string[] headers = { "【來源】庫", "來源表", "比對欄(如:日期)", "要同步之欄位", "➡️", "【目標】庫", "寫入目標表", "比對欄(如:年月)", "接收寫入之欄位", "同步方向" };
            for(int i=0; i<10; i++) {
                tlpSync.Controls.Add(new Label { Text = headers[i], TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) }, i, 0);
            }

            for (int i = 0; i < 5; i++)
            {
                var rowUi = new SyncRowUI();
                rowUi.CboSrcDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSrcTable = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSrcMatchCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSrcSyncCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };

                Label lblArrow = new Label { Text = "➡️", TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill };

                rowUi.CboTgtDb = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboTgtTable = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboTgtMatchCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboTgtSyncCol = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList }; 
                
                rowUi.CboSyncType = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
                rowUi.CboSyncType.Items.AddRange(new string[] { "", "單向同步", "雙向同步" });
                rowUi.CboSyncType.SelectedIndex = 0;

                BindSyncRowEvents(rowUi);

                tlpSync.Controls.Add(rowUi.CboSrcDb, 0, i+1);
                tlpSync.Controls.Add(rowUi.CboSrcTable, 1, i+1);
                tlpSync.Controls.Add(rowUi.CboSrcMatchCol, 2, i+1);
                tlpSync.Controls.Add(rowUi.CboSrcSyncCol, 3, i+1);
                tlpSync.Controls.Add(lblArrow, 4, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtDb, 5, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtTable, 6, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtMatchCol, 7, i+1);
                tlpSync.Controls.Add(rowUi.CboTgtSyncCol, 8, i+1);
                tlpSync.Controls.Add(rowUi.CboSyncType, 9, i+1);

                _syncRows.Add(rowUi);
            }

            boxSync.Controls.Add(tlpSync);

            Panel pnlSyncBottom = new Panel { Dock = DockStyle.Bottom, Height = 140, Padding = new Padding(15) };
            Button btnSaveSync = new Button { Text = "儲存新增的資料同步設定", Location = new Point(15, 10), Size = new Size(250, 45), BackColor = Color.Teal, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveSync.Click += BtnSaveSync_Click;
            
            // 🟢 新增：同步設定清單 按鈕
            Button btnShowSyncRulesList = new Button { Text = "同步設定清單", Location = new Point(290, 10), Size = new Size(180, 45), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnShowSyncRulesList.Click += BtnShowSyncRulesList_Click;

            Label lblSyncInfo = new Label { Text = "※ 僅需填入要啟用的列即可，空列系統自動忽略。若目標表無該接收欄位，系統會自動新增。\n※ 同步支援【雙向同步】(限1:1對應)，若比對欄位不一致(如 日期 vs 年月)，為防呆將強制設為單向聚合加總。", Location = new Point(15, 65), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };
            
            pnlSyncBottom.Controls.Add(btnSaveSync);
            pnlSyncBottom.Controls.Add(btnShowSyncRulesList);
            pnlSyncBottom.Controls.Add(lblSyncInfo);
            boxSync.Controls.Add(pnlSyncBottom);

            // 5. 刪除資料表區塊
            GroupBox boxDelete = new GroupBox { Text = "🔥 強制刪除整個資料表 (極度危險操作)", Dock = DockStyle.Top, Height = 230, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.Crimson, Padding = new Padding(15) };
            boxDelete.Margin = new Padding(0, 30, 0, 0);

            Label lblDelDesc = new Label { Text = "若資料表結構異常，您可於此將整張資料表永久刪除。刪除後重新點擊模組選單即可自動建立乾淨的空表。", AutoSize = true, Location = new Point(30, 45), ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };

            Label lblDelDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 100), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.Black };
            _cboDelDb = new ComboBox { Location = new Point(150, 98), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblDelTable = new Label { Text = "選擇資料表:", Location = new Point(360, 100), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.Black };
            _cboDelTable = new ComboBox { Location = new Point(480, 98), Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnExecuteDelete = new Button { Text = "⚠️ 執行永久刪除", Location = new Point(760, 95), Size = new Size(180, 40), BackColor = Color.Crimson, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnExecuteDelete.Click += BtnExecuteDelete_Click;

            boxDelete.Controls.AddRange(new Control[] { lblDelDesc, lblDelDb, _cboDelDb, lblDelTable, _cboDelTable, btnExecuteDelete });

            Panel spacer1 = new Panel { Dock = DockStyle.Top, Height = 30 };
            Panel spacer2 = new Panel { Dock = DockStyle.Top, Height = 30 };
            Panel spacer3 = new Panel { Dock = DockStyle.Top, Height = 30 };
            Panel spacer4 = new Panel { Dock = DockStyle.Top, Height = 30 };

            main.Controls.Add(boxDelete);
            main.Controls.Add(spacer4);
            main.Controls.Add(boxSync);
            main.Controls.Add(spacer3);
            main.Controls.Add(boxKeys);
            main.Controls.Add(spacer1);
            main.Controls.Add(boxBackup);
            main.Controls.Add(spacer2);
            main.Controls.Add(boxPath);

            // 🟢 加入空白選單項目，讓使用者可以取消選擇
            var blankDb = new ItemMap { EnName = "", ChName = "" };
            _cboDb.Items.Add(blankDb);
            _cboDelDb.Items.Add(blankDb);
            
            foreach (var sr in _syncRows) {
                sr.CboSrcDb.Items.Add(blankDb);
                sr.CboTgtDb.Items.Add(blankDb);
            }

            foreach (var kvp in _dbMap) {
                _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                _cboDelDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                
                foreach (var sr in _syncRows) {
                    sr.CboSrcDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                    sr.CboTgtDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
                }
            }
            
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
            _cboDelDb.SelectedIndexChanged += CboDelDb_SelectedIndexChanged;

            // 🟢 進入畫面時，預設選擇空白選項 (防重寫與同步都不會自動帶出舊資料)
            if (_cboDb.Items.Count > 0) _cboDb.SelectedIndex = 0;
            if (_cboDelDb.Items.Count > 0) _cboDelDb.SelectedIndex = 0;

            return main;
        }

        private void BindSyncRowEvents(SyncRowUI r)
        {
            Action checkSyncType = () => {
                string srcM = r.CboSrcMatchCol.Text;
                string tgtM = r.CboTgtMatchCol.Text;
                if (!string.IsNullOrEmpty(srcM) && !string.IsNullOrEmpty(tgtM)) {
                    if (srcM != tgtM || (srcM.Contains("日") && tgtM.Contains("月"))) {
                        r.CboSyncType.SelectedIndex = 1; // 單向同步
                        r.CboSyncType.Enabled = false;
                    } else {
                        r.CboSyncType.Enabled = true;
                    }
                }
            };

            r.CboSrcDb.SelectedIndexChanged += (s, e) => {
                r.CboSrcTable.Items.Clear(); r.CboSrcMatchCol.Items.Clear(); r.CboSrcSyncCol.Items.Clear();
                r.CboSrcTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
                
                if (r.CboSrcDb.SelectedItem != null) {
                    var map = (ItemMap)r.CboSrcDb.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && _dbMap.ContainsKey(map.EnName)) {
                        foreach (var tbl in _dbMap[map.EnName].Tables) r.CboSrcTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                    }
                }
            };
            r.CboSrcTable.SelectedIndexChanged += (s, e) => {
                r.CboSrcMatchCol.Items.Clear(); r.CboSrcSyncCol.Items.Clear();
                r.CboSrcMatchCol.Items.Add(""); r.CboSrcSyncCol.Items.Add("");
                
                if (r.CboSrcTable.SelectedItem != null) {
                    var map = (ItemMap)r.CboSrcDb.SelectedItem;
                    var tmap = (ItemMap)r.CboSrcTable.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && !string.IsNullOrEmpty(tmap.EnName)) {
                        var cols = DataManager.GetColumnNames(map.EnName, tmap.EnName);
                        foreach (var c in cols) if (c != "Id") { r.CboSrcMatchCol.Items.Add(c); r.CboSrcSyncCol.Items.Add(c); }
                    }
                }
            };

            r.CboTgtDb.SelectedIndexChanged += (s, e) => {
                r.CboTgtTable.Items.Clear(); r.CboTgtMatchCol.Items.Clear(); r.CboTgtSyncCol.Items.Clear();
                r.CboTgtTable.Items.Add(new ItemMap { EnName = "", ChName = "" });

                if (r.CboTgtDb.SelectedItem != null) {
                    var map = (ItemMap)r.CboTgtDb.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && _dbMap.ContainsKey(map.EnName)) {
                        foreach (var tbl in _dbMap[map.EnName].Tables) r.CboTgtTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                    }
                }
            };
            r.CboTgtTable.SelectedIndexChanged += (s, e) => {
                r.CboTgtMatchCol.Items.Clear(); r.CboTgtSyncCol.Items.Clear();
                r.CboTgtSyncCol.DropDownStyle = ComboBoxStyle.DropDown;
                r.CboTgtMatchCol.Items.Add(""); r.CboTgtSyncCol.Items.Add("");

                if (r.CboTgtTable.SelectedItem != null) {
                    var map = (ItemMap)r.CboTgtDb.SelectedItem;
                    var tmap = (ItemMap)r.CboTgtTable.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && !string.IsNullOrEmpty(tmap.EnName)) {
                        var cols = DataManager.GetColumnNames(map.EnName, tmap.EnName);
                        foreach (var c in cols) if (c != "Id") { r.CboTgtMatchCol.Items.Add(c); r.CboTgtSyncCol.Items.Add(c); }
                    }
                }
            };

            r.CboSrcMatchCol.TextChanged += (s, e) => checkSyncType();
            r.CboTgtMatchCol.TextChanged += (s, e) => checkSyncType();
        }

        // 🟢 防重寫設定清單 (視窗)
        private void BtnShowTableKeysList_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "防重寫設定清單", Size = new Size(800, 500), StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White })
            {
                Label lbl = new Label { Text = "已設定之資料表防重寫規則：", Dock = DockStyle.Top, Padding = new Padding(15), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10) };
                f.Controls.Add(flp);

                Action loadKeys = null;
                loadKeys = () => {
                    flp.Controls.Clear();
                    try {
                        string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT * FROM TableKeys", conn)) {
                                using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                            }
                        }

                        if (dt.Rows.Count == 0) {
                            flp.Controls.Add(new Label { Text = "目前沒有任何防重寫設定。", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) });
                            return;
                        }

                        foreach (DataRow row in dt.Rows) {
                            string db = row["DbName"].ToString();
                            string tb = row["TableName"].ToString();
                            string c1 = row["Col1"].ToString(); string c2 = row["Col2"].ToString();
                            string c3 = row["Col3"].ToString(); string c4 = row["Col4"].ToString();

                            Panel p = new Panel { Width = 740, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5) };
                            Label lTxt = new Label { Text = $"庫:[{db}] 表:[{tb}]  ➡️ 規則: {c1} | {c2} | {c3} | {c4}", AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F) };
                            
                            Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(680, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                            btnDel.FlatAppearance.BorderSize = 0;
                            btnDel.Click += (s, ev) => {
                                if (MessageBox.Show($"確定刪除 [{tb}] 的防重寫設定？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                                    if (AuthManager.VerifyAdmin()) {
                                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                                            conn.Open();
                                            using (var cmd = new SQLiteCommand("DELETE FROM TableKeys WHERE DbName=@DB AND TableName=@TB", conn)) {
                                                cmd.Parameters.AddWithValue("@DB", db); cmd.Parameters.AddWithValue("@TB", tb);
                                                cmd.ExecuteNonQuery();
                                            }
                                        }
                                        loadKeys();
                                    }
                                }
                            };

                            p.Controls.Add(lTxt);
                            p.Controls.Add(btnDel);
                            flp.Controls.Add(p);
                        }
                    } catch { }
                };

                loadKeys();
                f.ShowDialog();
            }
        }

        // 🟢 同步設定清單 (視窗)
        private void BtnShowSyncRulesList_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "同步設定清單", Size = new Size(800, 500), StartPosition = FormStartPosition.CenterParent, MaximizeBox = false, MinimizeBox = false, BackColor = Color.White })
            {
                Label lbl = new Label { Text = "已啟用之跨表同步規則清單：", Dock = DockStyle.Top, Padding = new Padding(15), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
                f.Controls.Add(lbl);

                FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10) };
                f.Controls.Add(flp);

                Action loadRules = null;
                loadRules = () => {
                    flp.Controls.Clear();
                    try {
                        string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT * FROM SyncRules", conn)) {
                                using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                            }
                        }

                        if (dt.Rows.Count == 0) {
                            flp.Controls.Add(new Label { Text = "目前沒有任何同步設定。", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) });
                            return;
                        }

                        foreach (DataRow row in dt.Rows) {
                            int id = Convert.ToInt32(row["Id"]);
                            string sDb = row["SrcDb"].ToString(); string sTb = row["SrcTable"].ToString(); string sSync = row["SrcSyncCol"].ToString();
                            string tDb = row["TgtDb"].ToString(); string tTb = row["TgtTable"].ToString(); string tSync = row["TgtSyncCol"].ToString();
                            string type = row.Table.Columns.Contains("SyncType") ? row["SyncType"].ToString() : "單向同步";

                            Panel p = new Panel { Width = 740, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5) };
                            Label lTxt = new Label { Text = $"【{type}】 {sDb}.{sTb}[{sSync}]  ➡️  {tDb}.{tTb}[{tSync}]", AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F) };
                            
                            Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(680, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                            btnDel.FlatAppearance.BorderSize = 0;
                            btnDel.Click += (s, ev) => {
                                if (MessageBox.Show($"確定刪除此同步規則？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                                    if (AuthManager.VerifyAdmin()) {
                                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                                            conn.Open();
                                            using (var cmd = new SQLiteCommand("DELETE FROM SyncRules WHERE Id=@Id", conn)) {
                                                cmd.Parameters.AddWithValue("@Id", id);
                                                cmd.ExecuteNonQuery();
                                            }
                                        }
                                        loadRules();
                                    }
                                }
                            };

                            p.Controls.Add(lTxt);
                            p.Controls.Add(btnDel);
                            flp.Controls.Add(p);
                        }
                    } catch { }
                };

                loadRules();
                f.ShowDialog();
            }
        }

        // 🟢 驗證隱藏選單密碼的輔助視窗
        private bool VerifyHiddenMenuPassword(string menuName)
        {
            using (Form p = new Form())
            {
                p.Width = 460; 
                p.Height = 220;
                p.Text = "隱藏選單安全驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;
                p.BackColor = Color.White;

                Label lbl = new Label() { Left = 30, Top = 30, Text = $"同步規則包含隱藏表單，請輸入【{menuName}】的密碼：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 250, Left = 30, Top = 70, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認驗證", DialogResult = DialogResult.OK, Left = 160, Top = 120, Width = 120, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog() == DialogResult.OK)
                {
                    string input = txt.Text.Trim();
                    string unlockedMenu = App_PasswordManager.CheckUnlockMenu(input);
                    if (unlockedMenu == menuName)
                    {
                        return true;
                    }
                    else
                    {
                        MessageBox.Show($"【{menuName}】密碼錯誤！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                return false; 
            }
        }

        private void BtnSaveSync_Click(object sender, EventArgs e)
        {
            if (!AuthManager.VerifyAdmin()) return;

            HashSet<string> requiredMenusToUnlock = new HashSet<string>();
            
            foreach (var r in _syncRows) 
            {
                if (r.CboSrcDb.SelectedItem == null || r.CboTgtDb.SelectedItem == null) continue;
                
                string srcDb = ((ItemMap)r.CboSrcDb.SelectedItem).EnName;
                string tgtDb = ((ItemMap)r.CboTgtDb.SelectedItem).EnName;
                
                if (string.IsNullOrEmpty(srcDb) || string.IsNullOrEmpty(tgtDb)) continue;

                if (srcDb == "Menu1DB" || tgtDb == "Menu1DB") requiredMenusToUnlock.Add("選單1");
                if (srcDb == "Menu2DB" || tgtDb == "Menu2DB") requiredMenusToUnlock.Add("選單2");
                if (srcDb == "Menu3DB" || tgtDb == "Menu3DB") requiredMenusToUnlock.Add("選單3");
                if (srcDb == "Menu4DB" || tgtDb == "Menu4DB") requiredMenusToUnlock.Add("選單4");
            }

            foreach (string menu in requiredMenusToUnlock)
            {
                if (!VerifyHiddenMenuPassword(menu))
                {
                    MessageBox.Show($"已取消儲存！因為您未通過【{menu}】的密碼驗證。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; 
                }
            }

            try {
                string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {

                        foreach (var r in _syncRows) {
                            if (r.CboSrcDb.SelectedItem == null || r.CboSrcTable.SelectedItem == null || string.IsNullOrWhiteSpace(r.CboSrcMatchCol.Text) || string.IsNullOrWhiteSpace(r.CboSrcSyncCol.Text) ||
                                r.CboTgtDb.SelectedItem == null || r.CboTgtTable.SelectedItem == null || string.IsNullOrWhiteSpace(r.CboTgtMatchCol.Text) || string.IsNullOrWhiteSpace(r.CboTgtSyncCol.Text))
                                continue;

                            string srcDb = ((ItemMap)r.CboSrcDb.SelectedItem).EnName;
                            string srcTbl = ((ItemMap)r.CboSrcTable.SelectedItem).EnName;
                            string tgtDb = ((ItemMap)r.CboTgtDb.SelectedItem).EnName;
                            string tgtTbl = ((ItemMap)r.CboTgtTable.SelectedItem).EnName;

                            if (string.IsNullOrEmpty(srcDb) || string.IsNullOrEmpty(tgtDb)) continue;

                            // 🟢 改為附加 (INSERT)，不要先 DELETE 全表，才不會覆蓋舊有的設定
                            string sql = "INSERT INTO SyncRules (SrcDb, SrcTable, SrcMatchCol, SrcSyncCol, TgtDb, TgtTable, TgtMatchCol, TgtSyncCol, SyncType) VALUES (@SD, @ST, @SMC, @SSC, @TD, @TT, @TMC, @TSC, @Type)";
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@SD", srcDb); cmd.Parameters.AddWithValue("@ST", srcTbl);
                                cmd.Parameters.AddWithValue("@SMC", r.CboSrcMatchCol.Text); cmd.Parameters.AddWithValue("@SSC", r.CboSrcSyncCol.Text);
                                cmd.Parameters.AddWithValue("@TD", tgtDb); cmd.Parameters.AddWithValue("@TT", tgtTbl);
                                cmd.Parameters.AddWithValue("@TMC", r.CboTgtMatchCol.Text); cmd.Parameters.AddWithValue("@TSC", r.CboTgtSyncCol.Text);
                                cmd.Parameters.AddWithValue("@Type", string.IsNullOrEmpty(r.CboSyncType.Text) ? "單向同步" : r.CboSyncType.Text);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                MessageBox.Show("新增的資料同步設定已成功儲存至清單！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // 儲存後將介面的 Combo 歸零
                foreach (var r in _syncRows) {
                    r.CboSrcDb.SelectedIndex = 0; r.CboTgtDb.SelectedIndex = 0;
                    r.CboSrcMatchCol.Text = ""; r.CboSrcSyncCol.Text = "";
                    r.CboTgtMatchCol.Text = ""; r.CboTgtSyncCol.Text = "";
                }

            } catch (Exception ex) {
                MessageBox.Show("儲存失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboTable == null || _cboCol1 == null) return;
                _cboTable.Items.Clear(); _cboCol1.Items.Clear(); _cboCol2.Items.Clear(); _cboCol3.Items.Clear(); _cboCol4.Items.Clear();
                _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });

                if (_cboDb.SelectedItem == null) return;
                
                var selectedDb = (ItemMap)_cboDb.SelectedItem;
                if (!string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                    foreach (var tbl in _dbMap[selectedDb.EnName].Tables) {
                        _cboTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                    }
                }
            } catch { }
        }

        private void CboDelDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboDelTable == null) return;
                _cboDelTable.Items.Clear();
                _cboDelTable.Items.Add(new ItemMap { EnName = "", ChName = "" });

                if (_cboDelDb.SelectedItem == null) return;
                
                var selectedDb = (ItemMap)_cboDelDb.SelectedItem;
                if (!string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                    foreach (var tbl in _dbMap[selectedDb.EnName].Tables) {
                        _cboDelTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                    }
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

                if (!string.IsNullOrEmpty(dbName) && !string.IsNullOrEmpty(tableName)) {
                    List<string> cols = DataManager.GetColumnNames(dbName, tableName);
                    foreach(var c in cols) if (c != "Id") { 
                        _cboCol1.Items.Add(c); _cboCol2.Items.Add(c); _cboCol3.Items.Add(c); _cboCol4.Items.Add(c);
                    }
                }
                
                // 🟢 預設為空，不主動載入已設定的值 (讓使用者自行去清單查看)
                _cboCol1.SelectedIndex = 0;
                _cboCol2.SelectedIndex = 0;
                _cboCol3.SelectedIndex = 0;
                _cboCol4.SelectedIndex = 0;
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
            
            if (string.IsNullOrEmpty(dbName) || string.IsNullOrEmpty(tableName)) {
                MessageBox.Show("請先選擇資料庫與資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string c1 = _cboCol1.SelectedItem?.ToString() ?? "";
            string c2 = _cboCol2.SelectedItem?.ToString() ?? "";
            string c3 = _cboCol3.SelectedItem?.ToString() ?? "";
            string c4 = _cboCol4.SelectedItem?.ToString() ?? "";

            DataManager.SaveTableKeys(dbName, tableName, c1, c2, c3, c4);
            MessageBox.Show($"【{chTableName}】 防重寫規則儲存成功！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // 儲存完後歸零
            _cboDb.SelectedIndex = 0;
        }

        private void BtnExecuteDelete_Click(object sender, EventArgs e)
        {
            if (_cboDelDb.SelectedItem == null || _cboDelTable.SelectedItem == null) {
                MessageBox.Show("請先選擇要刪除的資料庫與資料表！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string dbName = ((ItemMap)_cboDelDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboDelTable.SelectedItem).EnName;
            string chTableName = ((ItemMap)_cboDelTable.SelectedItem).ChName;

            if (string.IsNullOrEmpty(dbName) || string.IsNullOrEmpty(tableName)) {
                MessageBox.Show("請先選擇要刪除的資料庫與資料表！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            if (!AuthManager.VerifyTableDelete()) return;

            try {
                DataManager.DropTable(dbName, tableName);
                MessageBox.Show($"【{chTableName}】資料表已成功刪除。\n請從左側選單重新進入該模組以建立新結構。", "執行成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _cboDelDb.SelectedIndex = 0;
            } catch (Exception ex) {
                MessageBox.Show("刪除失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
