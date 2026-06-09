/// FILE: Safety_System/settings/App_DbConfig.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public partial class App_DbConfig : Form
    {
        // ================= 全域與共用 UI 控制項 =================
        private TextBox _txtPath, _txtAttachmentPath, _txtBackupPath;
        private NumericUpDown _numKeepCount, _numIntervalDays; 
        
        private ComboBox _cboDb, _cboTable, _cboCol1, _cboCol2, _cboCol3, _cboCol4;
        private ComboBox _cboDelDb, _cboDelTable;
        private ComboBox _cboAuditDb, _cboAuditTable;
        private DataGridView _dgvAudit;
        private CheckBox _chkShowDeletedLogs; 

        private ComboBox _cboFormulaDb, _cboFormulaTable, _cboFormulaTargetCol, _cboFormulaMatchCol;
        private DateTimePicker _dtpFormulaStart, _dtpFormulaEnd;
        private RichTextBox _rtbFormulaEditor;
        private FlowLayoutPanel _flpFormulasList;
        private int _currentFormulaEditId = 0; 

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

        private Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // ================= 靜態快取字典 =================
        public static Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> GetDbMapCache()
        {
            Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> map = new Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> {
                { "Safety", ("工安", new Dictionary<string, string> { 
                    { "NearMiss", "虛驚事件" }, { "SafetyInspection", "巡檢記錄" }, { "SafetyObservation", "安全觀察" }, 
                    { "TrafficInjury", "交通意外" }, { "WorkInjury", "工傷事件" }, { "MinorInjury", "輕傷事件" }, { "LaborInspection", "勞檢稽查缺失" } 
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
                    { "WaterVolume", "自來水用量統計" }, { "WaterUsageDaily", "自來水使用量" }, { "WaterPermitMaterial", "水污許可(原物料)" } 
                })},
                { "Waste", ("廢棄物", new Dictionary<string, string> { 
                    { "WasteMonthly", "廢棄物月表" }, { "Waste_IL", "複層月表" }, { "Waste_LM", "膠合月表" }, { "Waste_CR", "鍍板月表" }, 
                    { "Waste_T", "強化月表" }, { "Waste_GCTE", "切磨月表" }, { "Waste_ML", "物料月表" }, { "Waste_Water", "水站月表" },
                    { "WastePermitMaterial", "廢棄物污許可(原物料)" }, { "WastePermitProduct", "廢棄物污許可(產品)" }, { "WastePermitWaste", "廢棄物污許可(廢棄物)" }, { "WasteDisposalRecord", "廢棄物清運記錄" }
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
                { "ISO14001", ("ISO14001", new Dictionary<string, string> { { "TargetManagement", "目標管理" }, { "EnvInfoReceive", "環境資訊接收管制表" }, { "InternalComm", "內文聯絡書管制表" }, { "MailReceive", "郵件收文管制表" }, { "VisitorRecord", "來賓拜訪紀錄表" } })},
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
                        if (map.ContainsKey(dbName)) map[dbName].Tables[tableName] = tableName;
                    }
                }
            } catch { }
            return map;
        }

        public App_DbConfig()
        {
            CheckAndUpgradeFormulaTable();
        }

        private void CheckAndUpgradeFormulaTable()
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    var cols = new List<string>();
                    using (var cmd = new SQLiteCommand("PRAGMA table_info([ColumnFormulas])", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) cols.Add(reader["name"].ToString());
                    }

                    if (!cols.Contains("MatchCol")) {
                        string createSql = @"CREATE TABLE IF NOT EXISTS [ColumnFormulas_v2] (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT, DbName TEXT, TableName TEXT, 
                            TargetCol TEXT, MatchCol TEXT, StartDate TEXT, EndDate TEXT, Formula TEXT);";
                        using(var cmd = new SQLiteCommand(createSql, conn)) cmd.ExecuteNonQuery();
                        
                        string migrateSql = "INSERT INTO ColumnFormulas_v2 (DbName, TableName, TargetCol, MatchCol, StartDate, EndDate, Formula) " +
                                            "SELECT DbName, TableName, TargetCol, DateCol, '1900-01-01', '2099-12-31', Formula FROM ColumnFormulas";
                        using(var cmd = new SQLiteCommand(migrateSql, conn)) cmd.ExecuteNonQuery();

                        using(var cmd = new SQLiteCommand("DROP TABLE ColumnFormulas", conn)) cmd.ExecuteNonQuery();
                        using(var cmd = new SQLiteCommand("ALTER TABLE ColumnFormulas_v2 RENAME TO ColumnFormulas", conn)) cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        // ================= 核心視窗建置 =================
        public Control GetView()
        {
            _dbMap = GetDbMapCache();

            Panel mainPanel = new Panel { Dock = DockStyle.Fill };
            TabControl tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(15, 8) };

            // 呼叫各 Partial Class 內的 UI 建置方法
            TabPage tabSync = new TabPage("🔄 資料庫同步與聚合") { BackColor = Color.WhiteSmoke };
            BuildSyncTab(tabSync);

            TabPage tabFormula = new TabPage("🧮 欄位自動運算") { BackColor = Color.WhiteSmoke };
            BuildFormulaTab(tabFormula);

            TabPage tabKeys = new TabPage("🛡️ 寫入防呆") { BackColor = Color.WhiteSmoke };
            BuildKeysTab(tabKeys);

            TabPage tabDb = new TabPage("💾 管理與備份") { BackColor = Color.WhiteSmoke };
            BuildManagementTab(tabDb);

            TabPage tabAudit = new TabPage("🕵️ 軌跡查詢") { BackColor = Color.WhiteSmoke };
            BuildAuditTab(tabAudit);

            tabControl.TabPages.Add(tabSync);
            tabControl.TabPages.Add(tabFormula);
            tabControl.TabPages.Add(tabKeys);
            tabControl.TabPages.Add(tabDb);
            tabControl.TabPages.Add(tabAudit);
            
            mainPanel.Controls.Add(tabControl);

            InitializeSharedDropdowns();

            return mainPanel;
        }

        private void InitializeSharedDropdowns()
        {
            var blankDb = new ItemMap { EnName = "", ChName = "" };
            var dbItems = _dbMap.Select(kvp => new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName }).ToArray();

            _cboDb.Items.Add(blankDb); _cboDb.Items.AddRange(dbItems);
            _cboDelDb.Items.Add(blankDb); _cboDelDb.Items.AddRange(dbItems);
            _cboAuditDb.Items.Add(blankDb); _cboAuditDb.Items.AddRange(dbItems);
            _cboFormulaDb.Items.Add(blankDb); _cboFormulaDb.Items.AddRange(dbItems);
            
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
            if (_cboFormulaDb.Items.Count > 0) _cboFormulaDb.SelectedIndex = 0;
        }

        // ================= 共用事件邏輯 =================
        private void CboDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try {
                if (_cboTable == null || _cboCol1 == null) return;
                var selectedDb = (ItemMap)_cboDb.SelectedItem;

                if (selectedDb != null && selectedDb.EnName.StartsWith("Menu") && selectedDb.EnName.EndsWith("DB")) {
                    string menuName = selectedDb.EnName.Replace("DB", "").Replace("Menu", "選單");
                    if (!AuthManager.VerifyHiddenMenu(menuName)) {
                        _cboDb.SelectedIndex = 0; 
                        return;
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
                
                _cboCol1.SelectedIndex = 0; _cboCol2.SelectedIndex = 0; _cboCol3.SelectedIndex = 0; _cboCol4.SelectedIndex = 0;
            } catch { }
        }

        // 🟢 確保此方法存在於 App_DbConfig 主類別中，因為這是在多個分頁中共用的！
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
    }
}
