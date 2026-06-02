/// FILE: Safety_System/settings/App_DbConfig.cs ///
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
        private NumericUpDown _numIntervalDays; 
        
        private ComboBox _cboDb, _cboTable, _cboCol1, _cboCol2, _cboCol3, _cboCol4;
        private ComboBox _cboDelDb, _cboDelTable;

        // 稽核日誌專用變數
        private ComboBox _cboAuditDb, _cboAuditTable;
        private DataGridView _dgvAudit;
        private CheckBox _chkShowDeletedLogs; 

        private class SyncRowUI {
            public ComboBox CboSrcDb, CboSrcTable, CboSrcMatchCol, CboSrcSyncCol;
            public ComboBox CboTgtDb, CboTgtTable, CboTgtMatchCol, CboTgtSyncCol;
            public ComboBox CboSyncType; 
        }
        private List<SyncRowUI> _syncRows = new List<SyncRowUI>();

        public class ItemMap {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        public static Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> GetDbMapCache()
        {
            Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> map = new Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> {
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
                { "Purchase", ("日常作業", new Dictionary<string, string> { { "PurchaseData", "請購資料" } })},
                { "Menu1DB", ("選單1", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })},
                { "Menu2DB", ("選單2", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })},
                { "Menu3DB", ("選單3", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })},
                { "Menu4DB", ("選單4", new Dictionary<string, string> { { "WorkItems", "WorkItems" } })}
            };

            try {
                DataTable dtMenus = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
                if (dtMenus != null) {
                    foreach (DataRow r in dtMenus.Rows) {
                        string dbName = r["資料庫名"].ToString();
                        string tableName = r["資料表名"].ToString();
                        if (map.ContainsKey(dbName)) {
                            map[dbName].Tables[tableName] = tableName;
                        }
                    }
                }
            } catch { }
            return map;
        }

        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        public Control GetView()
        {
            _dbMap = GetDbMapCache();

            Panel mainPanel = new Panel { Dock = DockStyle.Fill };
            
            TabControl tabControl = new TabControl { 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 12F), 
                Padding = new Point(15, 8) 
            };

            // ==========================================
            // 分頁 1: 同步
            // ==========================================
            TabPage tabSync = new TabPage("🔄 資料同步");
            tabSync.BackColor = Color.WhiteSmoke;
            Panel pnlSync = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };
            
            GroupBox boxSync = new GroupBox { Text = "資料同步設定 (來源儲存時自動聚合計算至目標表)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };

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
            
            Button btnShowSyncRulesList = new Button { Text = "同步設定清單", Location = new Point(290, 10), Size = new Size(180, 45), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnShowSyncRulesList.Click += BtnShowSyncRulesList_Click;

            Label lblSyncInfo = new Label { Text = "※ 僅需填入要啟用的列即可，空列系統自動忽略。若目標表無該接收欄位，系統會自動新增。\n※ 同步支援【雙向同步】(限1:1對應)，若比對欄位不一致(如 日期 vs 年月)，為防呆將強制設為單向聚合加總。", Location = new Point(15, 65), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };
            
            pnlSyncBottom.Controls.Add(btnSaveSync);
            pnlSyncBottom.Controls.Add(btnShowSyncRulesList);
            pnlSyncBottom.Controls.Add(lblSyncInfo);
            boxSync.Controls.Add(pnlSyncBottom);

            pnlSync.Controls.Add(boxSync);
            tabSync.Controls.Add(pnlSync);

            // ==========================================
            // 分頁 2: 防重寫
            // ==========================================
            TabPage tabKeys = new TabPage("🛡️ 寫入防呆");
            tabKeys.BackColor = Color.WhiteSmoke;
            Panel pnlKeys = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

            GroupBox boxKeys = new GroupBox { Text = "資料表防重寫欄位設定 (空值則正常寫入不防呆)", Dock = DockStyle.Top, Height = 360, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };

            Label lblDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 50), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboDb = new ComboBox { Location = new Point(160, 48), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblTable = new Label { Text = "選擇資料表:", Location = new Point(420, 50), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable = new ComboBox { Location = new Point(540, 48), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol1 = new Label { Text = "判斷欄位一:", Location = new Point(30, 110), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol1 = new ComboBox { Location = new Point(160, 108), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol2 = new Label { Text = "判斷欄位二:", Location = new Point(420, 110), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol2 = new ComboBox { Location = new Point(540, 108), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol3 = new Label { Text = "判斷欄位三:", Location = new Point(30, 170), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol3 = new ComboBox { Location = new Point(160, 168), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol4 = new Label { Text = "判斷欄位四:", Location = new Point(420, 170), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol4 = new ComboBox { Location = new Point(540, 168), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnSaveKeys = new Button { Text = "儲存防重寫規則", Location = new Point(30, 240), Size = new Size(220, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveKeys.Click += BtnSaveKeys_Click;

            Button btnShowTableKeysList = new Button { Text = "防重寫設定清單", Location = new Point(280, 240), Size = new Size(200, 45), BackColor = Color.LightSlateGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnShowTableKeysList.Click += BtnShowTableKeysList_Click;

            boxKeys.Controls.AddRange(new Control[] { 
                lblDb, _cboDb, lblTable, _cboTable, 
                lblCol1, _cboCol1, lblCol2, _cboCol2, 
                lblCol3, _cboCol3, lblCol4, _cboCol4, 
                btnSaveKeys, btnShowTableKeysList 
            });

            pnlKeys.Controls.Add(boxKeys);
            tabKeys.Controls.Add(pnlKeys);

            // ==========================================
            // 分頁 3: 資料庫 (路徑、備份、刪除)
            // ==========================================
            TabPage tabDb = new TabPage("💾 資料庫管理");
            tabDb.BackColor = Color.WhiteSmoke;
            Panel pnlDb = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

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
                string authPrompt = "變更資料庫路徑需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (!AuthManager.VerifyAdmin(authPrompt)) return; 

                if (System.IO.Directory.Exists(_txtPath.Text)) {
                    DataManager.SetBasePath(_txtPath.Text);
                    MessageBox.Show("路徑已更新！後續系統存取皆會依此路徑。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("請選擇有效的資料夾路徑。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            boxPath.Controls.AddRange(new Control[] { _txtPath, btnBrowse, btnSavePath });

            GroupBox boxBackup = new GroupBox { Text = "資料庫備份設定 (背景自動執行)", Dock = DockStyle.Top, Height = 300, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
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

            Label lblB2 = new Label { Text = "保留舊備份份數:", Location = new Point(30, 105), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _numKeepCount = new NumericUpDown { Location = new Point(200, 103), Width = 80, Minimum = 1, Maximum = 365, Value = BackupManager.KeepCount, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblB3 = new Label { Text = "份 (建議保留 30 份，約一個月)", Location = new Point(290, 105), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };

            Label lblB4 = new Label { Text = "自動備份執行頻率:", Location = new Point(30, 160), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _numIntervalDays = new NumericUpDown { Location = new Point(200, 158), Width = 80, Minimum = 1, Maximum = 30, Value = BackupManager.BackupIntervalDays, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblB5 = new Label { Text = "天執行一次 (建議設為 1 天)", Location = new Point(290, 160), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) };

            Button btnSaveBackup = new Button { Text = "儲存備份設定", Location = new Point(30, 220), Size = new Size(220, 45), BackColor = Color.Sienna, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveBackup.Click += (s, e) => {
                string authPrompt = "修改備份設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (!AuthManager.VerifyAdmin(authPrompt)) return;

                BackupManager.SaveConfig(_txtBackupPath.Text, (int)_numKeepCount.Value, (int)_numIntervalDays.Value);
                MessageBox.Show("備份設定已儲存！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Button btnManualBackup = new Button { Text = "立即執行手動熱備份", Location = new Point(300, 220), Size = new Size(220, 45), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnManualBackup.Click += (s, e) => {
                BackupManager.ExecuteBackup();
                MessageBox.Show("熱備份(Hot Backup)執行完成！\n不會影響目前操作中的使用者。", "備份成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            boxBackup.Controls.AddRange(new Control[] { lblB1, _txtBackupPath, btnBrowseBackup, lblB2, _numKeepCount, lblB3, lblB4, _numIntervalDays, lblB5, btnSaveBackup, btnManualBackup });

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

            pnlDb.Controls.Add(boxDelete);
            pnlDb.Controls.Add(spacer2);
            pnlDb.Controls.Add(boxBackup);
            pnlDb.Controls.Add(spacer1);
            pnlDb.Controls.Add(boxPath);

            tabDb.Controls.Add(pnlDb);

            // ==========================================
            // 分頁 4: 軌跡查詢
            // ==========================================
            TabPage tabAudit = new TabPage("🕵️ 軌跡查詢");
            tabAudit.BackColor = Color.WhiteSmoke;
            Panel pnlAudit = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

            GroupBox boxAudit = new GroupBox { Text = "操作軌跡追蹤 (查閱最後修改人與時間)", Dock = DockStyle.Top, Height = 650, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Padding = new Padding(15) };

            Label lblAuditDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.Black };
            _cboAuditDb = new ComboBox { Location = new Point(140, 58), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblAuditTable = new Label { Text = "選擇資料表:", Location = new Point(350, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.Black };
            _cboAuditTable = new ComboBox { Location = new Point(460, 58), Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            _chkShowDeletedLogs = new CheckBox { Text = "☑️ 僅查詢該表「被刪除的資料」軌跡", Location = new Point(740, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.Crimson, Cursor = Cursors.Hand };

            Button btnSearchAudit = new Button { Text = "🔍 查詢操作紀錄", Location = new Point(740, 110), Size = new Size(180, 35), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnSearchAudit.Click += BtnSearchAudit_Click;

            _dgvAudit = new DataGridView { 
                Location = new Point(30, 170), 
                Size = new Size(1000, 440),
                BackgroundColor = Color.WhiteSmoke, 
                AllowUserToAddRows = false, 
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, 
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Microsoft JhengHei UI", 11F)
            };
            
            // 🟢 同步開啟雙緩衝防閃爍
            typeof(DataGridView).InvokeMember("DoubleBuffered", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, 
                null, _dgvAudit, new object[] { true });

            boxAudit.Controls.AddRange(new Control[] { lblAuditDb, _cboAuditDb, lblAuditTable, _cboAuditTable, _chkShowDeletedLogs, btnSearchAudit, _dgvAudit });

            pnlAudit.Controls.Add(boxAudit);
            tabAudit.Controls.Add(pnlAudit);

            // ==========================================
            // 將分頁加入 TabControl
            // ==========================================
            tabControl.TabPages.Add(tabSync);
            tabControl.TabPages.Add(tabKeys);
            tabControl.TabPages.Add(tabDb);
            tabControl.TabPages.Add(tabAudit);
            
            mainPanel.Controls.Add(tabControl);

            // 🟢 優化：改用陣列批次匯入，避免逐項 Add 觸發數百次畫面重繪
            var blankDb = new ItemMap { EnName = "", ChName = "" };
            var dbItems = _dbMap.Select(kvp => new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName }).ToArray();

            _cboDb.Items.Add(blankDb); _cboDb.Items.AddRange(dbItems);
            _cboDelDb.Items.Add(blankDb); _cboDelDb.Items.AddRange(dbItems);
            _cboAuditDb.Items.Add(blankDb); _cboAuditDb.Items.AddRange(dbItems);
            
            foreach (var sr in _syncRows) {
                sr.CboSrcDb.Items.Add(blankDb); sr.CboSrcDb.Items.AddRange(dbItems);
                sr.CboTgtDb.Items.Add(blankDb); sr.CboTgtDb.Items.AddRange(dbItems);
            }
            
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
            _cboDelDb.SelectedIndexChanged += CboDelDb_SelectedIndexChanged;
            _cboAuditDb.SelectedIndexChanged += CboAuditDb_SelectedIndexChanged;

            if (_cboDb.Items.Count > 0) _cboDb.SelectedIndex = 0;
            if (_cboDelDb.Items.Count > 0) _cboDelDb.SelectedIndex = 0;
            if (_cboAuditDb.Items.Count > 0) _cboAuditDb.SelectedIndex = 0;

            return mainPanel;
        }

        private void BtnSearchAudit_Click(object sender, EventArgs e)
        {
            if (_cboAuditDb.SelectedItem == null || _cboAuditTable.SelectedItem == null) {
                MessageBox.Show("請先選擇要查詢的資料庫與資料表！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string dbName = ((ItemMap)_cboAuditDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboAuditTable.SelectedItem).EnName;

            if (string.IsNullOrEmpty(dbName) || string.IsNullOrEmpty(tableName)) {
                MessageBox.Show("請先選擇要查詢的資料庫與資料表！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            try 
            {
                if (_chkShowDeletedLogs.Checked)
                {
                    DataTable dtDel = new DataTable();
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        string sql = "SELECT RecordId AS [原系統流水號(Id)], DeletedBy AS [執行刪除者], DeletedTime AS [刪除時間] FROM System_DeleteLogs WHERE DbName=@DB AND TableName=@TB ORDER BY DeletedTime DESC";
                        using (var cmd = new SQLiteCommand(sql, conn)) {
                            cmd.Parameters.AddWithValue("@DB", dbName);
                            cmd.Parameters.AddWithValue("@TB", tableName);
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dtDel);
                        }
                    }

                    if (dtDel.Rows.Count > 0) {
                        _dgvAudit.DataSource = dtDel;
                        _dgvAudit.Columns["執行刪除者"].DefaultCellStyle.ForeColor = Color.Crimson;
                        _dgvAudit.Columns["執行刪除者"].DefaultCellStyle.Font = new Font(_dgvAudit.Font, FontStyle.Bold);
                    } else {
                        MessageBox.Show("該資料表目前沒有任何被刪除的紀錄。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        _dgvAudit.DataSource = null;
                    }
                }
                else
                {
                    DataTable dt = DataManager.GetTableData(dbName, tableName, "", "", "");
                    
                    if (dt != null && dt.Rows.Count > 0) 
                    {
                        DataTable dtAudit = new DataTable();
                        dtAudit.Columns.Add("系統流水號(Id)", typeof(int));
                        
                        if (dt.Columns.Contains("日期")) dtAudit.Columns.Add("發生日期", typeof(string));
                        else if (dt.Columns.Contains("年月")) dtAudit.Columns.Add("發生年月", typeof(string));
                        
                        if (dt.Columns.Contains("法規名稱")) dtAudit.Columns.Add("法規名稱", typeof(string));
                        else if (dt.Columns.Contains("化學物質名稱")) dtAudit.Columns.Add("化學物質名稱", typeof(string));

                        dtAudit.Columns.Add("最後修改人", typeof(string));
                        dtAudit.Columns.Add("最後修改時間", typeof(string));

                        foreach (DataRow r in dt.Rows) 
                        {
                            DataRow newRow = dtAudit.NewRow();
                            newRow["系統流水號(Id)"] = r["Id"];
                            
                            if (dt.Columns.Contains("日期")) newRow["發生日期"] = r["日期"];
                            else if (dt.Columns.Contains("年月")) newRow["發生年月"] = r["年月"];
                            
                            if (dt.Columns.Contains("法規名稱")) newRow["法規名稱"] = r["法規名稱"];
                            else if (dt.Columns.Contains("化學物質名稱")) newRow["化學物質名稱"] = r["化學物質名稱"];

                            newRow["最後修改人"] = dt.Columns.Contains("最後修改人") ? r["最後修改人"]?.ToString() : "無紀錄";
                            newRow["最後修改時間"] = dt.Columns.Contains("修改時間") ? r["修改時間"]?.ToString() : "無紀錄";

                            dtAudit.Rows.Add(newRow);
                        }

                        dtAudit.DefaultView.Sort = "最後修改時間 DESC";
                        _dgvAudit.DataSource = dtAudit.DefaultView.ToTable();
                        _dgvAudit.Columns["最後修改人"].DefaultCellStyle.ForeColor = Color.DarkSlateBlue;
                        _dgvAudit.Columns["最後修改人"].DefaultCellStyle.Font = new Font(_dgvAudit.Font, FontStyle.Bold);
                    }
                    else
                    {
                        MessageBox.Show("該資料表目前沒有任何現存資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        _dgvAudit.DataSource = null;
                    }
                }
            } 
            catch (Exception ex) 
            {
                MessageBox.Show("查詢失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CboAuditDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboAuditTable == null) return;
                _cboAuditTable.Items.Clear();
                _cboAuditTable.Items.Add(new ItemMap { EnName = "", ChName = "" });

                if (_cboAuditDb.SelectedItem == null) return;
                
                var selectedDb = (ItemMap)_cboAuditDb.SelectedItem;
                if (!string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                    var tbItems = _dbMap[selectedDb.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                    _cboAuditTable.Items.AddRange(tbItems);
                }
            } catch { }
        }

        private void BindSyncRowEvents(SyncRowUI r)
        {
            Action checkSyncType = () => {
                string srcM = r.CboSrcMatchCol.Text;
                string tgtM = r.CboTgtMatchCol.Text;
                if (!string.IsNullOrEmpty(srcM) && !string.IsNullOrEmpty(tgtM)) {
                    if (srcM != tgtM || (srcM.Contains("日") && tgtM.Contains("月"))) {
                        r.CboSyncType.SelectedIndex = 1; 
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
                        var tbItems = _dbMap[map.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                        r.CboSrcTable.Items.AddRange(tbItems);
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
                        var cols = DataManager.GetColumnNames(map.EnName, tmap.EnName).Where(c => c != "Id").ToArray();
                        r.CboSrcMatchCol.Items.AddRange(cols); r.CboSrcSyncCol.Items.AddRange(cols);
                    }
                }
            };

            r.CboTgtDb.SelectedIndexChanged += (s, e) => {
                r.CboTgtTable.Items.Clear(); r.CboTgtMatchCol.Items.Clear(); r.CboTgtSyncCol.Items.Clear();
                r.CboTgtTable.Items.Add(new ItemMap { EnName = "", ChName = "" });

                if (r.CboTgtDb.SelectedItem != null) {
                    var map = (ItemMap)r.CboTgtDb.SelectedItem;
                    if (!string.IsNullOrEmpty(map.EnName) && _dbMap.ContainsKey(map.EnName)) {
                        var tbItems = _dbMap[map.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                        r.CboTgtTable.Items.AddRange(tbItems);
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
                        var cols = DataManager.GetColumnNames(map.EnName, tmap.EnName).Where(c => c != "Id").ToArray();
                        r.CboTgtMatchCol.Items.AddRange(cols); r.CboTgtSyncCol.Items.AddRange(cols);
                    }
                }
            };

            r.CboSrcMatchCol.TextChanged += (s, e) => checkSyncType();
            r.CboTgtMatchCol.TextChanged += (s, e) => checkSyncType();
        }

        private void BtnShowTableKeysList_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { 
                Text = "防重寫設定清單", 
                Size = new Size(900, 550), 
                StartPosition = FormStartPosition.CenterParent, 
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true, 
                MinimizeBox = false, 
                BackColor = Color.White 
            })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                Label lbl = new Label { 
                    Text = "已設定之資料表防重寫規則 (每個資料表僅限一組)：", 
                    Dock = DockStyle.Fill, 
                    Padding = new Padding(15, 15, 15, 10), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    AutoSize = true
                };
                tlp.Controls.Add(lbl, 0, 0);

                FlowLayoutPanel flp = new FlowLayoutPanel { 
                    Dock = DockStyle.Fill, 
                    AutoScroll = true, 
                    FlowDirection = FlowDirection.TopDown, 
                    WrapContents = false, 
                    Padding = new Padding(10) 
                };
                tlp.Controls.Add(flp, 0, 1);
                f.Controls.Add(tlp);

                flp.Resize += (s, ev) => {
                    foreach (Control c in flp.Controls) {
                        if (c is Panel pnl) {
                            Label l = pnl.Controls.OfType<Label>().FirstOrDefault();
                            int minW = l != null ? TextRenderer.MeasureText(l.Text, l.Font).Width + 100 : 840;
                            pnl.Width = Math.Max(flp.ClientSize.Width - 25, minW);
                        }
                    }
                };

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

                            string text = $"庫:[{db}] 表:[{tb}]  ➡️ 規則: {c1} | {c2} | {c3} | {c4}";
                            Label lTxt = new Label { Text = text, AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F) };
                            
                            int reqW = TextRenderer.MeasureText(text, lTxt.Font).Width + 100;
                            int panelW = Math.Max(flp.ClientSize.Width - 25, reqW);

                            Panel p = new Panel { Width = panelW, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5) };
                            
                            Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(panelW - 60, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                            btnDel.FlatAppearance.BorderSize = 0;
                            btnDel.Click += (s, ev) => {
                                if (MessageBox.Show($"確定刪除 [{tb}] 的防重寫設定？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                                    string authPrompt = "刪除防重寫設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                                    if (AuthManager.VerifyAdmin(authPrompt)) {
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

        private void BtnShowSyncRulesList_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { 
                Text = "同步設定清單", 
                Size = new Size(900, 550), 
                StartPosition = FormStartPosition.CenterParent, 
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true, 
                MinimizeBox = false, 
                BackColor = Color.White 
            })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                Label lbl = new Label { 
                    Text = "已啟用之跨表同步規則清單：", 
                    Dock = DockStyle.Fill, 
                    Padding = new Padding(15, 15, 15, 10), 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    AutoSize = true
                };
                tlp.Controls.Add(lbl, 0, 0);

                FlowLayoutPanel flp = new FlowLayoutPanel { 
                    Dock = DockStyle.Fill, 
                    AutoScroll = true, 
                    FlowDirection = FlowDirection.TopDown, 
                    WrapContents = false, 
                    Padding = new Padding(10) 
                };
                tlp.Controls.Add(flp, 0, 1);
                f.Controls.Add(tlp);

                flp.Resize += (s, ev) => {
                    foreach (Control c in flp.Controls) {
                        if (c is Panel pnl) {
                            Label l = pnl.Controls.OfType<Label>().FirstOrDefault();
                            int minW = l != null ? TextRenderer.MeasureText(l.Text, l.Font).Width + 100 : 840;
                            pnl.Width = Math.Max(flp.ClientSize.Width - 25, minW);
                        }
                    }
                };

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

                            string text = $"【{type}】 {sDb}.{sTb}[{sSync}]  ➡️  {tDb}.{tTb}[{tSync}]";
                            Label lTxt = new Label { Text = text, AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F) };
                            
                            int reqW = TextRenderer.MeasureText(text, lTxt.Font).Width + 100;
                            int panelW = Math.Max(flp.ClientSize.Width - 25, reqW);

                            Panel p = new Panel { Width = panelW, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5) };
                            
                            Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(panelW - 60, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                            btnDel.FlatAppearance.BorderSize = 0;
                            btnDel.Click += (s, ev) => {
                                if (MessageBox.Show($"確定刪除此同步規則？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                                    string authPrompt = "刪除資料同步設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                                    if (AuthManager.VerifyAdmin(authPrompt)) {
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

        private void BtnSaveSync_Click(object sender, EventArgs e)
        {
            string authPrompt = "新增資料同步設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

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
                // 🟢 統一呼叫 AuthManager 進行密碼驗證
                if (!AuthManager.VerifyHiddenMenu(menu))
                {
                    MessageBox.Show($"已取消儲存！因為您未通過【{menu}】的密碼驗證。", "權限不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; 
                }
            }

            try {
                string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                
                HashSet<string> existingRules = new HashSet<string>();
                using (var connCheck = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                    connCheck.Open();
                    using (var cmdCheck = new SQLiteCommand("SELECT * FROM SyncRules", connCheck))
                    using (var reader = cmdCheck.ExecuteReader()) {
                        while (reader.Read()) {
                            string key = $"{reader["SrcDb"]}|{reader["SrcTable"]}|{reader["SrcMatchCol"]}|{reader["SrcSyncCol"]}|{reader["TgtDb"]}|{reader["TgtTable"]}|{reader["TgtMatchCol"]}|{reader["TgtSyncCol"]}";
                            existingRules.Add(key);
                        }
                    }
                }

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

                            string currentKey = $"{srcDb}|{srcTbl}|{r.CboSrcMatchCol.Text}|{r.CboSrcSyncCol.Text}|{tgtDb}|{tgtTbl}|{r.CboTgtMatchCol.Text}|{r.CboTgtSyncCol.Text}";
                            
                            if (existingRules.Contains(currentKey)) {
                                MessageBox.Show($"偵測到重複的同步設定！\n\n來源表：【{((ItemMap)r.CboSrcTable.SelectedItem).ChName}】\n目標表：【{((ItemMap)r.CboTgtTable.SelectedItem).ChName}】\n\n您設定的相同條件與欄位已經存在於系統中。\n為了防止資料庫產生雙重計算錯誤，請勿重複新增！", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                trans.Rollback();
                                return; 
                            }
                            
                            existingRules.Add(currentKey);

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

                var selectedDb = (ItemMap)_cboDb.SelectedItem;

                if (selectedDb != null && selectedDb.EnName.StartsWith("Menu") && selectedDb.EnName.EndsWith("DB"))
                {
                    string menuName = "";
                    if (selectedDb.EnName == "Menu1DB") menuName = "選單1";
                    if (selectedDb.EnName == "Menu2DB") menuName = "選單2";
                    if (selectedDb.EnName == "Menu3DB") menuName = "選單3";
                    if (selectedDb.EnName == "Menu4DB") menuName = "選單4";

                    if (!string.IsNullOrEmpty(menuName))
                    {
                        // 🟢 統一呼叫 AuthManager 進行密碼驗證
                        if (!AuthManager.VerifyHiddenMenu(menuName))
                        {
                            _cboDb.SelectedIndex = 0; 
                            return;
                        }
                    }
                }

                _cboTable.Items.Clear(); _cboCol1.Items.Clear(); _cboCol2.Items.Clear(); _cboCol3.Items.Clear(); _cboCol4.Items.Clear();
                _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });

                if (selectedDb != null && !string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                    var tbItems = _dbMap[selectedDb.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                    _cboTable.Items.AddRange(tbItems);
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
                    var tbItems = _dbMap[selectedDb.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                    _cboDelTable.Items.AddRange(tbItems);
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
                    var cols = DataManager.GetColumnNames(dbName, tableName).Where(c => c != "Id").ToArray();
                    _cboCol1.Items.AddRange(cols); _cboCol2.Items.AddRange(cols); _cboCol3.Items.AddRange(cols); _cboCol4.Items.AddRange(cols);
                }
                
                _cboCol1.SelectedIndex = 0;
                _cboCol2.SelectedIndex = 0;
                _cboCol3.SelectedIndex = 0;
                _cboCol4.SelectedIndex = 0;
            } catch { }
        }

        private void BtnSaveKeys_Click(object sender, EventArgs e)
        {
            string authPrompt = "修改防重寫設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return; 

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
