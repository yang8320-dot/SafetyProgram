/// FILE: Safety_System/settings/App_DropdownManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public class App_DropdownManager : Form
    {
        private TabControl _tabControl;

        // ================= Tab 1: 單選連動下拉選單控制項 =================
        private ComboBox _cboDb, _cboTable;
        private TextBox[] _txtOptions;
        private ComboBox[] _cboCols;
        private ComboBox[] _cboParentVals;
        private Button _btnSave, _btnExport, _btnImport, _btnClearAll, _btnClearDb;
        
        private Dictionary<string, bool> _configuredDbs = new Dictionary<string, bool>();
        private Dictionary<string, bool> _configuredTables = new Dictionary<string, bool>();
        private Dictionary<string, bool> _configuredCols = new Dictionary<string, bool>();

        private bool _isRevertingDb = false;
        private bool _isRevertingCol = false;
        private bool _isRevertingMultiCol = false; // 🟢 用於控制 Tab 2 複選框的防呆還原狀態

        // ================= Tab 2: 複選組合文字控制項 =================
        private ComboBox _cboDbMulti, _cboTableMulti, _cboColMulti;
        private TextBox _txtOptionsMulti;
        private Button _btnSaveMulti, _btnDelMulti;
        private FlowLayoutPanel _flpMultiConfigured;
        private Panel _selectedMultiItemPanel = null; 

        private class ItemMap 
        {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        private readonly Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        // 🟢 全域快取
        public static Dictionary<string, string[]> DropdownCache = new Dictionary<string, string[]>();
        public static Dictionary<string, string[]> MultiSelectCache = new Dictionary<string, string[]>();

        public App_DropdownManager()
        {
            try 
            {
                string sql = "CREATE TABLE IF NOT EXISTS [MultiSelectConfigs] (Id INTEGER PRIMARY KEY AUTOINCREMENT, TableName TEXT, ColName TEXT, Options TEXT, UNIQUE(TableName, ColName));";
                DataManager.InitTable("SystemConfig", "MultiSelectConfigs", sql);

                _dbMap = App_DbConfig.GetDbMapCache();
                RefreshConfiguredCache();
                InitializeComponent();
                LoadDropdownConfigs();
                LoadMultiSelectConfigs();
                RefreshMultiConfiguredList();
            } 
            catch (Exception ex) 
            {
                MessageBox.Show($"初始化連動選單管理介面時發生嚴重錯誤：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 🟢 智慧讀取：若實體表尚未建立，強制從 SchemaMap 讀取預設欄位結構
        private List<string> GetColumnsSafe(string dbName, string tbName)
        {
            var cols = DataManager.GetColumnNames(dbName, tbName);
            if (cols == null || cols.Count <= 1) 
            {
                cols = new List<string>();
                if (TableSchemaManager.SchemaMap.ContainsKey(tbName)) 
                {
                    string schema = TableSchemaManager.SchemaMap[tbName];
                    var parts = schema.Split(',');
                    cols.Add("Id"); 
                    foreach(var p in parts) 
                    {
                        int start = p.IndexOf('[');
                        int end = p.IndexOf(']');
                        if (start >= 0 && end > start) 
                        {
                            cols.Add(p.Substring(start + 1, end - start - 1));
                        }
                    }
                }
            }
            return cols;
        }

        // 🟢 輔助方法：檢查該欄位是否已存在於「單選下拉」設定中
        private bool IsColumnInDropdownCache(string tbName, string colName)
        {
            string prefix = $"{tbName}|{colName}|";
            foreach (var key in DropdownCache.Keys) {
                if (key.StartsWith(prefix)) return true;
            }
            return false;
        }

        private void RefreshConfiguredCache()
        {
            _configuredDbs.Clear();
            _configuredTables.Clear();
            _configuredCols.Clear();
            try 
            {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                {
                    conn.Open();
                    using(var cmd = new SQLiteCommand("SELECT TableName, ColName, ParentColName FROM DropdownConfigs", conn))
                    using(var reader = cmd.ExecuteReader()) 
                    {
                        while(reader.Read()) 
                        {
                            string tb = reader["TableName"].ToString();
                            string col = reader["ColName"].ToString();
                            string pCol = reader["ParentColName"].ToString();

                            bool isLinked = !string.IsNullOrEmpty(pCol);

                            if (!_configuredTables.ContainsKey(tb)) _configuredTables[tb] = false;
                            if (isLinked) _configuredTables[tb] = true;

                            string colKey = $"{tb}_{col}";
                            if (!_configuredCols.ContainsKey(colKey)) _configuredCols[colKey] = false;
                            if (isLinked) _configuredCols[colKey] = true;

                            if (isLinked) 
                            {
                                string pColKey = $"{tb}_{pCol}";
                                _configuredCols[pColKey] = true; 
                            }
                        }
                    }
                }
                
                if (_dbMap != null) 
                {
                    foreach(var kvp in _dbMap) 
                    {
                        string dbName = kvp.Key;
                        foreach(var tb in kvp.Value.Tables.Keys) 
                        {
                            if (_configuredTables.ContainsKey(tb)) 
                            {
                                if (!_configuredDbs.ContainsKey(dbName)) _configuredDbs[dbName] = false;
                                if (_configuredTables[tb]) _configuredDbs[dbName] = true;
                            }
                        }
                    }
                }
            } 
            catch { }
        }

        private void InitializeComponent()
        {
            this.Text = "下拉選單與組合文字(複選)管理中心";
            this.Size = new Size(1650, 900);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.BackColor = Color.WhiteSmoke;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(20, 10) };

            TabPage tabSingle = new TabPage("一、單選下拉與多層連動設定");
            tabSingle.BackColor = Color.WhiteSmoke;
            BuildTabSingle(tabSingle);

            TabPage tabMulti = new TabPage("二、組合文字 (複選) 設定");
            tabMulti.BackColor = Color.WhiteSmoke;
            BuildTabMulti(tabMulti);

            _tabControl.TabPages.Add(tabSingle);
            _tabControl.TabPages.Add(tabMulti);
            this.Controls.Add(_tabControl);
        }

        // =========================================================
        // Tab 1 UI 構建與邏輯 (單選下拉選單)
        // =========================================================
        private void BuildTabSingle(TabPage page)
        {
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 95, BackColor = Color.White, Padding = new Padding(20) };
            pnlBottom.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _btnSave = new Button { Text = "💾 儲存並套用當前設定", Width = 260, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSave.Click += BtnSave_Click;

            _btnClearAll = new Button { Text = "🗑️ 一鍵清除畫面上設定", Width = 260, Height = 50, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnClearAll.Click += BtnClearAll_Click;

            _btnClearDb = new Button { Text = "💣 清除所有資料庫設定", Width = 260, Height = 50, BackColor = Color.Crimson, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnClearDb.Click += BtnClearDb_Click;

            Label lblHint = new Label { Text = "※ 已設定過且「僅單層獨立」的項目，以【亮藍色粗體】標示。\n※ 已設定過且具備「多層連動關係」的項目，以【紅色粗體】標示。\n※ 選項內容的排列順序，即為系統表單中下拉選單顯示的順序。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0) };

            pnlBottom.Controls.Add(lblHint);
            
            FlowLayoutPanel flpBtnBottom = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false };
            flpBtnBottom.Controls.Add(_btnSave);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 }); 
            flpBtnBottom.Controls.Add(_btnClearAll);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 }); 
            flpBtnBottom.Controls.Add(_btnClearDb);
            
            pnlBottom.Controls.Add(flpBtnBottom);
            page.Controls.Add(pnlBottom); 

            Panel pnlTop = new Panel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Padding = new Padding(20) };
            pnlTop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            FlowLayoutPanel flpTopMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            Label lblTitle = new Label { Text = "🔧 單選下拉與多層連動設定", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            
            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            
            Label lblDb = new Label { Text = "選擇資料庫：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboDb = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 30, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboDb.DrawItem += CboDb_DrawItem;

            Label lblTable = new Label { Text = "選擇資料表：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboTable = new ComboBox { Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 40, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboTable.DrawItem += CboTable_DrawItem;

            _btnExport = new Button { Text = "📤 匯出 Excel", Size = new Size(150, 40), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 15, 0) };
            _btnImport = new Button { Text = "📥 匯入 Excel", Size = new Size(150, 40), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 0) };

            flpControls.Controls.AddRange(new Control[] { lblDb, _cboDb, lblTable, _cboTable, _btnExport, _btnImport });
            flpTopMain.Controls.Add(lblTitle);
            flpTopMain.Controls.Add(flpControls);
            pnlTop.Controls.Add(flpTopMain);
            page.Controls.Add(pnlTop); 

            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, Padding = new Padding(10, 15, 10, 15) };
            for(int i = 0; i < 4; i++) tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

            _cboCols = new ComboBox[4];
            _cboParentVals = new ComboBox[4];
            _txtOptions = new TextBox[4];

            string[] headers = { "第一層 (主選項)", "第二層 (依第一層連動)", "第三層 (依第二層連動)", "第四層 (依第三層連動)" };

            for (int i = 0; i < 4; i++)
            {
                Panel pColBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(3, 0, 3, 0), BackColor = Color.White, Padding = new Padding(15) };
                pColBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pColBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

                TableLayoutPanel pColInner = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7 };
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.AutoSize)); 
                pColInner.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 

                Label lHeader = new Label { Text = headers[i], Font = new Font("Microsoft JhengHei UI", 15F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0,0,0,15) };
                
                Label lCol = new Label { Text = "綁定資料表欄位：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0,0,0,5) };
                _cboCols[i] = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0,0,0,15), DrawMode = DrawMode.OwnerDrawFixed };
                int currentIndex = i; 
                _cboCols[i].DrawItem += (s, e) => CboCols_DrawItem(s, e, currentIndex);

                Label lParent = new Label { Text = "觸發條件 (父層選擇值)：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0,0,0,5) };
                _cboParentVals[i] = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0,0,0,15) };
                
                // 🟢 加入資料萃取按鈕，文字精簡為「導入資料」
                FlowLayoutPanel flpOptHeader = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0,0,0,5) };
                Label lOpt = new Label { Text = "下拉選項內容：", Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Margin = new Padding(0,5,5,0) };
                Button btnExtract = new Button { Text = "導入資料", Size = new Size(100, 30), BackColor = Color.LightSlateGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnExtract.FlatAppearance.BorderSize = 0;
                btnExtract.Click += (s, e) => ExtractDataFromDB(currentIndex, false);

                flpOptHeader.Controls.Add(lOpt);
                flpOptHeader.Controls.Add(btnExtract);

                _txtOptions[i] = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, WordWrap = false, Font = new Font("Microsoft JhengHei UI", 12F) };

                if (i == 0) {
                    lParent.Visible = false;
                    _cboParentVals[i].Visible = false;
                }

                pColInner.Controls.Add(lHeader, 0, 0);
                pColInner.Controls.Add(lCol, 0, 1);
                pColInner.Controls.Add(_cboCols[i], 0, 2);
                pColInner.Controls.Add(lParent, 0, 3);
                pColInner.Controls.Add(_cboParentVals[i], 0, 4);
                pColInner.Controls.Add(flpOptHeader, 0, 5);
                pColInner.Controls.Add(_txtOptions[i], 0, 6);

                int colIndex = i;
                _cboCols[colIndex].SelectedIndexChanged += (s, e) => HandleColSelectionChanged(colIndex);
                if (i > 0) _cboParentVals[colIndex].SelectedIndexChanged += (s, e) => HandleParentValChanged(colIndex);

                pColBorder.Controls.Add(pColInner);
                tlpMain.Controls.Add(pColBorder, i, 0);
            }

            page.Controls.Add(tlpMain); 
            tlpMain.BringToFront();     

            _btnExport.Click += BtnExport_Click;
            _btnImport.Click += BtnImport_Click;

            _cboDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_dbMap != null) 
            {
                foreach (var kvp in _dbMap) _cboDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });
            }
            
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;
        }

        // =========================================================
        // Tab 2 UI 構建與邏輯 (組合文字/複選)
        // =========================================================
        private void BuildTabMulti(TabPage page)
        {
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 95, BackColor = Color.White, Padding = new Padding(20) };
            pnlBottom.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlBottom.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            _btnSaveMulti = new Button { Text = "💾 儲存組合文字設定", Width = 230, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnSaveMulti.Click += BtnSaveMulti_Click;

            _btnDelMulti = new Button { Text = "🗑️ 刪除此欄位設定", Width = 230, Height = 50, BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            _btnDelMulti.Click += BtnDelMulti_Click;

            Button btnClearMultiDb = new Button { Text = "💣 清除所有資料庫設定", Width = 260, Height = 50, BackColor = Color.Crimson, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnClearMultiDb.Click += BtnClearMultiDb_Click;

            Label lblHintMulti = new Label { Text = "※ 已設定組合文字的項目，以【紅色粗體】標示。\n※ 組合文字(複選)設定後，該欄位於資料表中將變為強制唯讀，點擊後會彈出複選視窗。\n※ 匯入 Excel 格式：資料表名稱、欄位名稱、選項內容(逗號分隔)。", Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F), Padding = new Padding(0) };
            
            pnlBottom.Controls.Add(lblHintMulti);

            FlowLayoutPanel flpBtnBottom = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, WrapContents = false };
            flpBtnBottom.Controls.Add(_btnSaveMulti);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 });
            flpBtnBottom.Controls.Add(_btnDelMulti);
            flpBtnBottom.Controls.Add(new Panel { Width = 15, Height = 10 });
            flpBtnBottom.Controls.Add(btnClearMultiDb);
            
            pnlBottom.Controls.Add(flpBtnBottom);
            page.Controls.Add(pnlBottom);

            Panel pnlTop = new Panel { Dock = DockStyle.Top, AutoSize = true, MinimumSize = new Size(0, 110), BackColor = Color.White, Padding = new Padding(20) };
            pnlTop.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            FlowLayoutPanel flpTopMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            Label lblTitle = new Label { Text = "☑️ 組合文字 (複選) 設定區", Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), ForeColor = Color.DarkCyan, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };

            FlowLayoutPanel flpControls = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };

            Label l1 = new Label { Text = "選擇資料庫：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboDbMulti = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 20, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboDbMulti.DrawItem += CboDbMulti_DrawItem;
            
            _cboDbMulti.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_dbMap != null) { foreach (var kvp in _dbMap) _cboDbMulti.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName }); }

            Label l2 = new Label { Text = "選擇資料表：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboTableMulti = new ComboBox { Width = 260, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 20, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboTableMulti.DrawItem += CboTableMulti_DrawItem;

            Label l3 = new Label { Text = "指定目標欄位：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold), Margin = new Padding(0, 8, 5, 0) };
            _cboColMulti = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 4, 30, 0), DrawMode = DrawMode.OwnerDrawFixed };
            _cboColMulti.DrawItem += CboColMulti_DrawItem;

            Button btnExportMulti = new Button { Text = "📤 匯出 Excel", Size = new Size(150, 40), BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 15, 0) };
            btnExportMulti.Click += BtnExportMulti_Click;

            Button btnImportMulti = new Button { Text = "📥 匯入 Excel", Size = new Size(150, 40), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };
            btnImportMulti.Click += BtnImportMulti_Click;

            flpControls.Controls.AddRange(new Control[] { l1, _cboDbMulti, l2, _cboTableMulti, l3, _cboColMulti, btnExportMulti, btnImportMulti });
            flpTopMain.Controls.Add(lblTitle);
            flpTopMain.Controls.Add(flpControls);
            pnlTop.Controls.Add(flpTopMain);
            page.Controls.Add(pnlTop);

            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(10, 15, 10, 15) };
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            // 左側輸入框區塊
            Panel pnlLeftBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.White, Padding = new Padding(15) };
            pnlLeftBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlLeftBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
            
            // 🟢 加入資料萃取按鈕 (精簡文字為 導入資料)
            FlowLayoutPanel flpMultiOptHeader = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 10) };
            Label l4 = new Label { Text = "自訂核取選項：", Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, AutoSize = true, Margin = new Padding(0, 5, 10, 0) };
            Button btnExtractMulti = new Button { Text = "導入資料", Size = new Size(100, 30), BackColor = Color.LightSlateGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnExtractMulti.FlatAppearance.BorderSize = 0;
            btnExtractMulti.Click += (s, e) => ExtractDataFromDB(0, true);
            
            flpMultiOptHeader.Controls.Add(l4);
            flpMultiOptHeader.Controls.Add(btnExtractMulti);

            _txtOptionsMulti = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Microsoft JhengHei UI", 13F) };
            
            pnlLeftBorder.Controls.Add(_txtOptionsMulti);
            pnlLeftBorder.Controls.Add(flpMultiOptHeader); // 加在 Top

            // 右側清單區塊
            Panel pnlRightBorder = new Panel { Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.White, Padding = new Padding(15) };
            pnlRightBorder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlRightBorder.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);

            Label l5 = new Label { Text = "已設定之組合文字清單：", Dock = DockStyle.Top, Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Margin = new Padding(0, 0, 0, 10), Height = 30 };
            _flpMultiConfigured = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(5) };
            
            pnlRightBorder.Controls.Add(_flpMultiConfigured);
            pnlRightBorder.Controls.Add(l5);

            tlpMain.Controls.Add(pnlLeftBorder, 0, 0);
            tlpMain.Controls.Add(pnlRightBorder, 1, 0);

            page.Controls.Add(tlpMain);
            tlpMain.BringToFront();

            // 事件綁定
            _cboDbMulti.SelectedIndexChanged += (s, e) => {
                _cboTableMulti.Items.Clear(); _cboTableMulti.Items.Add(new ItemMap { EnName = "", ChName = "" }); _cboColMulti.Items.Clear(); _txtOptionsMulti.Clear();
                var db = _cboDbMulti.SelectedItem as ItemMap;
                if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                    foreach (var tbl in _dbMap[db.EnName].Tables) _cboTableMulti.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            };
            
            _cboTableMulti.SelectedIndexChanged += (s, e) => {
                _cboColMulti.Items.Clear(); _txtOptionsMulti.Clear();
                var db = _cboDbMulti.SelectedItem as ItemMap; var tb = _cboTableMulti.SelectedItem as ItemMap;
                if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                    // 智慧讀取：防呆讀不到空表的欄位
                    var cols = GetColumnsSafe(db.EnName, tb.EnName).Where(c => c != "Id" && c != "附件檔案");
                    foreach (var c in cols) _cboColMulti.Items.Add(c);
                }
            };

            // 🟢 防呆攔截：若選擇了已在單選下拉設定過的欄位，則阻擋
            _cboColMulti.SelectedIndexChanged += (s, e) => {
                if (_isRevertingMultiCol) return;

                _txtOptionsMulti.Clear();
                var tb = _cboTableMulti.SelectedItem as ItemMap;
                if (tb != null && _cboColMulti.SelectedItem != null) {
                    string colName = _cboColMulti.SelectedItem.ToString();
                    
                    if (IsColumnInDropdownCache(tb.EnName, colName)) {
                        MessageBox.Show($"此欄位【{colName}】已在「一、單選下拉與多層連動」中設定過！\n為避免系統判斷異常，同一欄位不可同時設定單選與複選。\n\n若要設定為組合文字，請先至第一分頁刪除該欄位的設定。", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _isRevertingMultiCol = true;
                        _cboColMulti.SelectedIndex = -1; // 取消選取
                        _isRevertingMultiCol = false;
                        return;
                    }

                    string key = $"{tb.EnName}|{colName}";
                    if (MultiSelectCache.ContainsKey(key)) {
                        _txtOptionsMulti.Text = string.Join(Environment.NewLine, MultiSelectCache[key]);
                    }
                }
            };
        }

        // =========================================================
        // Tab 2: 選單紅字加粗重繪邏輯 (OwnerDraw)
        // =========================================================
        private void CboDbMulti_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboDbMulti.Items[e.Index] as ItemMap;
            bool isConfig = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName)) {
                foreach(var k in MultiSelectCache.Keys) {
                    string tb = k.Split('|')[0];
                    if (_dbMap.ContainsKey(item.EnName) && _dbMap[item.EnName].Tables.ContainsKey(tb)) { isConfig = true; break; }
                }
            }
            DrawComboBoxItemMulti(_cboDbMulti, e, isConfig);
        }

        private void CboTableMulti_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboTableMulti.Items[e.Index] as ItemMap;
            bool isConfig = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName)) {
                foreach(var k in MultiSelectCache.Keys) {
                    if (k.StartsWith(item.EnName + "|")) { isConfig = true; break; }
                }
            }
            DrawComboBoxItemMulti(_cboTableMulti, e, isConfig);
        }

        private void CboColMulti_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            string col = _cboColMulti.Items[e.Index].ToString();
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            bool isConfig = false;
            if (tb != null && !string.IsNullOrEmpty(tb.EnName) && !string.IsNullOrEmpty(col)) {
                string key = $"{tb.EnName}|{col}";
                if (MultiSelectCache.ContainsKey(key)) isConfig = true;
            }
            DrawComboBoxItemMulti(_cboColMulti, e, isConfig);
        }

        private void DrawComboBoxItemMulti(ComboBox cbo, DrawItemEventArgs e, bool isConfigured) {
            if (e.Index < 0) return;
            
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected) 
                e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds);
            else 
                e.Graphics.FillRectangle(Brushes.White, e.Bounds);
            
            string text = cbo.Items[e.Index].ToString();
            Brush textBrush = Brushes.Black;
            Font currentFont = e.Font;
            
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected) {
                textBrush = Brushes.White;
                if (isConfigured) currentFont = new Font(e.Font, FontStyle.Bold);
            } else if (isConfigured) {
                textBrush = Brushes.Crimson; 
                currentFont = new Font(e.Font, FontStyle.Bold);
            }
            
            e.Graphics.DrawString(text, currentFont, textBrush, new RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height));
            e.DrawFocusRectangle();
        }

        // =========================================================
        // 🟢 全新功能：一鍵從現有資料表中萃取不重複文字
        // =========================================================
        private void ExtractDataFromDB(int colIndex, bool isMulti)
        {
            string db = isMulti ? ((ItemMap)_cboDbMulti.SelectedItem)?.EnName : ((ItemMap)_cboDb.SelectedItem)?.EnName;
            string tb = isMulti ? ((ItemMap)_cboTableMulti.SelectedItem)?.EnName : ((ItemMap)_cboTable.SelectedItem)?.EnName;
            string col = isMulti ? _cboColMulti.SelectedItem?.ToString() : _cboCols[colIndex].Text;

            if (string.IsNullOrEmpty(db) || string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) {
                MessageBox.Show("請先正確選擇資料庫、資料表與欄位！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try {
                DataTable dt = DataManager.GetTableData(db, tb, "", "", "");
                if (dt == null || dt.Rows.Count == 0) {
                    MessageBox.Show("此資料表目前尚無資料可供萃取。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                HashSet<string> uniqueVals = new HashSet<string>();
                foreach(DataRow r in dt.Rows) {
                    if (dt.Columns.Contains(col)) {
                        string v = r[col]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(v)) {
                            if (isMulti) {
                                // 複選文字會以換行符或逗號儲存，自動拆解
                                var parts = v.Split(new[] { '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach(var p in parts) {
                                    string cleanP = p.Trim();
                                    if (!string.IsNullOrEmpty(cleanP)) uniqueVals.Add(cleanP);
                                }
                            } else {
                                uniqueVals.Add(v);
                            }
                        }
                    }
                }

                if (uniqueVals.Count > 0) {
                    TextBox targetTxt = isMulti ? _txtOptionsMulti : _txtOptions[colIndex];
                    var currentLines = targetTxt.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
                    
                    int addedCount = 0;
                    foreach(var v in uniqueVals) {
                        if (!currentLines.Contains(v)) {
                            currentLines.Add(v);
                            addedCount++;
                        }
                    }
                    targetTxt.Text = string.Join(Environment.NewLine, currentLines);
                    MessageBox.Show($"導入完成！共新增 {addedCount} 筆不重複資料至文字框中。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("導入完成，但在該欄位沒有發現有效資料。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } catch (Exception ex) {
                MessageBox.Show("導入失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================
        // Tab 1 事件邏輯
        // =========================================================
        private void CboDb_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboDb.Items[e.Index] as ItemMap;
            bool isConfig = false;
            bool isLinked = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName) && _configuredDbs.ContainsKey(item.EnName)) {
                isConfig = true; isLinked = _configuredDbs[item.EnName];
            }
            DrawComboBoxItem(_cboDb, e, isConfig, isLinked);
        }

        private void CboTable_DrawItem(object sender, DrawItemEventArgs e) {
            if (e.Index < 0) return;
            var item = _cboTable.Items[e.Index] as ItemMap;
            bool isConfig = false;
            bool isLinked = false;
            if (item != null && !string.IsNullOrEmpty(item.EnName) && _configuredTables.ContainsKey(item.EnName)) {
                isConfig = true; isLinked = _configuredTables[item.EnName];
            }
            DrawComboBoxItem(_cboTable, e, isConfig, isLinked);
        }

        private void CboCols_DrawItem(object sender, DrawItemEventArgs e, int colIndex) {
            if (e.Index < 0) return;
            string colName = _cboCols[colIndex].Items[e.Index].ToString();
            string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName ?? "";
            bool isConfig = false;
            bool isLinked = false;
            string key = $"{tbName}_{colName}";
            if (!string.IsNullOrEmpty(colName) && _configuredCols.ContainsKey(key)) {
                isConfig = true; isLinked = _configuredCols[key];
            }
            DrawComboBoxItem(_cboCols[colIndex], e, isConfig, isLinked);
        }

        private void DrawComboBoxItem(ComboBox cbo, DrawItemEventArgs e, bool isConfigured, bool isLinked) {
            if (e.Index < 0) return;

            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected) 
                e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds);
            else 
                e.Graphics.FillRectangle(Brushes.White, e.Bounds);
            
            string text = cbo.Items[e.Index].ToString();
            Brush textBrush = Brushes.Black;
            
            using (Font boldFont = new Font(e.Font, FontStyle.Bold)) {
                Font currentFont = e.Font;
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected) {
                    textBrush = Brushes.White;
                    if (isConfigured) currentFont = boldFont; 
                } 
                else if (isConfigured) {
                    textBrush = isLinked ? Brushes.Crimson : Brushes.DodgerBlue; 
                    currentFont = boldFont;         
                }
                e.Graphics.DrawString(text, currentFont, textBrush, new RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height));
            }
            e.DrawFocusRectangle();
        }

        private void BtnClearAll_Click(object sender, EventArgs e) {
            if (MessageBox.Show("確定要清空畫面上所有的選項與連動設定嗎？\n(注意：尚未按下儲存前，資料庫內的設定並不會被刪除。)", "確認清除畫面", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) {
                ClearAllEditors();
            }
        }

        private void BtnClearDb_Click(object sender, EventArgs e) {
            if (MessageBox.Show("【極度危險】\n您確定要永久刪除資料庫中「所有」的單選下拉與連動設定嗎？\n此操作無法復原！", "永久刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes) {
                string authPrompt = "清除設定資料需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (!AuthManager.VerifyAdmin(authPrompt)) return;

                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM DropdownConfigs", conn)) { cmd.ExecuteNonQuery(); }
                    }
                    MessageBox.Show("資料庫設定已全部清空！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ClearAllEditors();
                    RefreshConfiguredCache();
                    LoadDropdownConfigs();
                    
                    _cboDb.SelectedIndex = 0; 
                    _cboDb.Invalidate();
                    _cboTable.Invalidate();
                    foreach(var c in _cboCols) c.Invalidate();
                } catch (Exception ex) {
                    MessageBox.Show($"清除失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool VerifyHiddenMenuPassword(string menuName) {
            using (Form p = new Form()) {
                p.Width = 460; 
                p.Height = 220;
                p.Text = "隱藏選單安全驗證";
                p.StartPosition = FormStartPosition.CenterParent;
                p.FormBorderStyle = FormBorderStyle.FixedDialog;
                p.MaximizeBox = false; 
                p.MinimizeBox = false;
                p.BackColor = Color.White;

                Label lbl = new Label() { Left = 30, Top = 30, Text = $"請輸入【{menuName}】的解鎖密碼以繼續設定：", AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox txt = new TextBox { PasswordChar = '*', Width = 250, Left = 30, Top = 70, Font = new Font("Microsoft JhengHei UI", 14F) };
                Button btn = new Button { Text = "確認驗證", DialogResult = DialogResult.OK, Left = 160, Top = 120, Width = 120, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };

                p.Controls.Add(lbl); 
                p.Controls.Add(txt); 
                p.Controls.Add(btn);
                p.AcceptButton = btn;

                if (p.ShowDialog(this) == DialogResult.OK) {
                    string input = txt.Text.Trim();
                    string unlockedMenu = App_PasswordManager.CheckUnlockMenu(input);
                    if (unlockedMenu == menuName) return true;
                    
                    MessageBox.Show($"【{menuName}】密碼錯誤！", "驗證失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false; 
            }
        }

        private void CboDb_SelectedIndexChanged(object sender, EventArgs e) {
            if (_isRevertingDb) return;
            var selectedDb = _cboDb.SelectedItem as ItemMap;
            if (selectedDb != null && selectedDb.EnName.StartsWith("Menu") && selectedDb.EnName.EndsWith("DB")) {
                string menuName = "";
                if (selectedDb.EnName == "Menu1DB") menuName = "選單1";
                if (selectedDb.EnName == "Menu2DB") menuName = "選單2";
                if (selectedDb.EnName == "Menu3DB") menuName = "選單3";
                if (selectedDb.EnName == "Menu4DB") menuName = "選單4";

                if (!string.IsNullOrEmpty(menuName)) {
                    if (!VerifyHiddenMenuPassword(menuName)) {
                        _isRevertingDb = true;
                        _cboDb.SelectedIndex = 0; 
                        _isRevertingDb = false;
                        return;
                    }
                }
            }

            _cboTable.Items.Clear();
            _cboTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
            ClearAllEditors();

            if (selectedDb != null && !string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                foreach (var tbl in _dbMap[selectedDb.EnName].Tables) {
                    _cboTable.Items.Add(new ItemMap { EnName = tbl.Key, ChName = tbl.Value });
                }
            }
            if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
        }

        private void CboTable_SelectedIndexChanged(object sender, EventArgs e) {
            ClearAllEditors();
            if (_cboDb.SelectedItem is ItemMap dbMap && _cboTable.SelectedItem is ItemMap tbMap && !string.IsNullOrEmpty(dbMap.EnName) && !string.IsNullOrEmpty(tbMap.EnName)) {
                // 🟢 智慧讀取欄位
                var cols = GetColumnsSafe(dbMap.EnName, tbMap.EnName);
                foreach (var cbo in _cboCols) {
                    cbo.Items.Clear();
                    cbo.Items.Add("");
                    foreach (var c in cols) if (c != "Id" && c != "附件檔案" && c != "備註") cbo.Items.Add(c);
                }
            }
        }

        private void ClearAllEditors() {
            _isRevertingCol = true;
            for (int i = 0; i < 4; i++) {
                if (_cboCols[i].Items.Count > 0) _cboCols[i].SelectedIndex = 0;
                if (i > 0) { _cboParentVals[i].Items.Clear(); _cboParentVals[i].Items.Add(""); }
                _txtOptions[i].Clear();
            }
            _isRevertingCol = false;
        }

        // 🟢 防呆攔截：若選擇了已在組合文字(複選)設定過的欄位，則阻擋
        private void HandleColSelectionChanged(int colIndex) {
            if (_isRevertingCol) return;
            string selectedCol = _cboCols[colIndex].Text;
            string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;
            
            if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(selectedCol)) {
                string multiKey = $"{tbName}|{selectedCol}";
                if (MultiSelectCache.ContainsKey(multiKey)) {
                    MessageBox.Show($"此欄位【{selectedCol}】已在「二、組合文字 (複選)」中設定過！\n為避免系統判斷異常，同一欄位不可同時設定單選與複選。\n\n若要設定為單選連動，請先至第二分頁刪除該欄位的設定。", "防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _isRevertingCol = true;
                    _cboCols[colIndex].SelectedIndex = 0; // 回復空白
                    _isRevertingCol = false;
                    return;
                }

                for (int i = 0; i < 4; i++) {
                    if (i != colIndex && _cboCols[i].Text == selectedCol) {
                        MessageBox.Show("此欄位已在其他層級被設定，為防止系統錯亂，請勿重複選擇！", "重複選擇防呆", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _isRevertingCol = true;
                        _cboCols[colIndex].SelectedIndex = 0; 
                        _isRevertingCol = false;
                        return;
                    }
                }
            }

            try {
                if (colIndex == 0) {
                    if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(selectedCol)) {
                        LoadOptionsToTextBox(tbName, selectedCol, "", "", _txtOptions[0]);
                        UpdateChildParentVals(1, _txtOptions[0].Text);
                    } else {
                        _txtOptions[0].Clear();
                        UpdateChildParentVals(1, "");
                    }
                }
            } catch { }
        }

        private void HandleParentValChanged(int colIndex) {
            try {
                if (colIndex <= 0 || colIndex >= 4) return;
                
                string colName = _cboCols[colIndex].Text;
                string parentVal = _cboParentVals[colIndex].Text;
                string parentCol = _cboCols[colIndex - 1].Text;
                string tbName = ((ItemMap)_cboTable.SelectedItem)?.EnName;

                if (!string.IsNullOrEmpty(tbName) && !string.IsNullOrEmpty(colName)) {
                    LoadOptionsToTextBox(tbName, colName, parentCol, parentVal, _txtOptions[colIndex]);
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, _txtOptions[colIndex].Text);
                } else {
                    _txtOptions[colIndex].Clear();
                    if (colIndex < 3) UpdateChildParentVals(colIndex + 1, "");
                }
            } catch { }
        }

        private void UpdateChildParentVals(int childIndex, string parentOptionsText) {
            try {
                if (childIndex <= 0 || childIndex >= 4) return;
                
                var opts = parentOptionsText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                string currentVal = _cboParentVals[childIndex].Text;
                
                _cboParentVals[childIndex].Items.Clear();
                _cboParentVals[childIndex].Items.Add("");
                foreach (var o in opts) _cboParentVals[childIndex].Items.Add(o.Trim());

                if (!string.IsNullOrEmpty(currentVal) && _cboParentVals[childIndex].Items.Contains(currentVal))
                    _cboParentVals[childIndex].Text = currentVal;
                else
                    _cboParentVals[childIndex].SelectedIndex = 0;
            } catch { }
        }

        private void LoadOptionsToTextBox(string tableName, string colName, string parentColName, string parentVal, TextBox txt) {
            string opts = GetDropdownOptionsFromDB(tableName, colName, parentColName, parentVal);
            txt.Text = opts.Replace(",", Environment.NewLine);
        }

        private string GetDropdownOptionsFromDB(string tableName, string colName, string parentColName, string parentVal) {
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = "SELECT Options FROM DropdownConfigs WHERE TableName=@T AND ColName=@C AND IFNULL(ParentColName,'')=@PC AND IFNULL(ParentValue,'')=@PV";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@T", tableName);
                        cmd.Parameters.AddWithValue("@C", colName);
                        cmd.Parameters.AddWithValue("@PC", parentColName);
                        cmd.Parameters.AddWithValue("@PV", parentVal);
                        var res = cmd.ExecuteScalar();
                        return res != null ? res.ToString() : "";
                    }
                }
            } catch { return ""; }
        }

        private void BtnSave_Click(object sender, EventArgs e) {
            if (_cboTable.SelectedItem == null || string.IsNullOrEmpty(((ItemMap)_cboTable.SelectedItem).EnName)) {
                MessageBox.Show("請先選擇資料表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string tbName = ((ItemMap)_cboTable.SelectedItem).EnName;

            // 🟢 強制儲存前雙重檢驗防呆
            for (int i = 0; i < 4; i++) {
                string colName = _cboCols[i].Text;
                if (!string.IsNullOrEmpty(colName) && MultiSelectCache.ContainsKey($"{tbName}|{colName}")) {
                    MessageBox.Show($"欄位【{colName}】已設定為「組合文字(複選)」，為避免系統錯亂，無法將其儲存為單選連動！", "儲存攔截", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }

            string authPrompt = "儲存下拉選單設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        for (int i = 0; i < 4; i++) {
                            string colName = _cboCols[i].Text;
                            if (string.IsNullOrEmpty(colName)) continue;

                            string parentCol = i > 0 ? _cboCols[i-1].Text : "";
                            string parentVal = i > 0 ? _cboParentVals[i].Text : "";
                            
                            var optsArray = _txtOptions[i].Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToArray();
                            string optsStr = string.Join(",", optsArray);

                            string sql = @"INSERT INTO DropdownConfigs (TableName, ColName, ParentColName, ParentValue, Options) 
                                           VALUES (@T, @C, @PC, @PV, @Opt) 
                                           ON CONFLICT(TableName, ColName, ParentColName, ParentValue) DO UPDATE SET Options=@Opt";
                            
                            using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                cmd.Parameters.AddWithValue("@T", tbName); cmd.Parameters.AddWithValue("@C", colName);
                                cmd.Parameters.AddWithValue("@PC", parentCol); cmd.Parameters.AddWithValue("@PV", parentVal);
                                cmd.Parameters.AddWithValue("@Opt", optsStr); cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                MessageBox.Show("選項設定已儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshConfiguredCache(); 
                _cboDb.Invalidate(); _cboTable.Invalidate(); foreach(var c in _cboCols) c.Invalidate(); 
                LoadDropdownConfigs(); 
                if (!string.IsNullOrEmpty(_txtOptions[0].Text)) UpdateChildParentVals(1, _txtOptions[0].Text);
            } catch (Exception ex) {
                MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e) {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統下拉選單設定_" + DateTime.Now.ToString("yyyyMMdd") }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT TableName AS [資料表名稱], ColName AS [欄位名稱], ParentColName AS [父層欄位], ParentValue AS [觸發條件], Options AS [選項內容(逗號分隔)] FROM DropdownConfigs", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }
                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("下拉選單設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("匯出成功！請直接在 Excel 中編輯後匯入。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e) {
            string authPrompt = "匯入下拉選單設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的設定檔" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string tb = ws.Cells[r, 1].Text.Trim(); string col = ws.Cells[r, 2].Text.Trim();
                                        string pCol = ws.Cells[r, 3].Text.Trim(); string pVal = ws.Cells[r, 4].Text.Trim(); string opt = ws.Cells[r, 5].Text.Trim();
                                        if (string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) continue;

                                        string sql = @"INSERT INTO DropdownConfigs (TableName, ColName, ParentColName, ParentValue, Options) VALUES (@T, @C, @PC, @PV, @Opt) ON CONFLICT(TableName, ColName, ParentColName, ParentValue) DO UPDATE SET Options=@Opt";
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@T", tb); cmd.Parameters.AddWithValue("@C", col);
                                            cmd.Parameters.AddWithValue("@PC", pCol); cmd.Parameters.AddWithValue("@PV", pVal); cmd.Parameters.AddWithValue("@Opt", opt);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("下拉選單設定已批次匯入並【自動存檔覆寫】成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        RefreshConfiguredCache(); LoadDropdownConfigs();
                        _cboDb.SelectedIndex = 0; _cboDb.Invalidate(); _cboTable.Invalidate(); foreach(var c in _cboCols) c.Invalidate();
                    } catch (Exception ex) { MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        public static void LoadDropdownConfigs() {
            DropdownCache.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM DropdownConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string tb = reader["TableName"].ToString(); string col = reader["ColName"].ToString();
                            string pCol = reader["ParentColName"].ToString(); string pVal = reader["ParentValue"].ToString(); string opts = reader["Options"].ToString();
                            string key = $"{tb}|{col}|{pCol}|{pVal}";
                            var arr = opts.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
                            DropdownCache[key] = arr;
                        }
                    }
                }
            } catch { }
        }

        public static string[] GetOptions(string tableName, string colName, string parentColName = "", string parentValue = "") {
            string key = $"{tableName}|{colName}|{parentColName}|{parentValue}";
            if (DropdownCache.ContainsKey(key)) return DropdownCache[key];
            return null;
        }

        public static string[] GetAllOptionsForColumn(string tableName, string colName) {
            HashSet<string> allOpts = new HashSet<string> { "" };
            foreach(var kvp in DropdownCache) {
                var parts = kvp.Key.Split('|');
                if(parts.Length == 4 && parts[0] == tableName && parts[1] == colName) foreach(var opt in kvp.Value) allOpts.Add(opt);
            }
            return allOpts.ToArray();
        }

        // =========================================================
        // Tab 2 事件邏輯 (組合文字/複選)
        // =========================================================
        private void BtnSaveMulti_Click(object sender, EventArgs e)
        {
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            if (tb == null || _cboColMulti.SelectedItem == null || string.IsNullOrWhiteSpace(_txtOptionsMulti.Text)) {
                MessageBox.Show("請確認資料表、欄位與選項內容皆已填寫！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string tblName = tb.EnName;
            string colName = _cboColMulti.SelectedItem.ToString();
            
            // 🟢 強制儲存前雙重檢驗防呆
            if (IsColumnInDropdownCache(tblName, colName)) {
                MessageBox.Show($"欄位【{colName}】已在「單選下拉與多層連動」設定中！\n為避免系統錯亂，不可同時設定單選與複選，無法儲存！", "儲存攔截", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            string authPrompt = "儲存組合文字設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            var optsArray = _txtOptionsMulti.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToArray();
            string optsStr = string.Join(",", optsArray);

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = @"INSERT INTO MultiSelectConfigs (TableName, ColName, Options) VALUES (@T, @C, @Opt) ON CONFLICT(TableName, ColName) DO UPDATE SET Options=@Opt";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@T", tblName);
                        cmd.Parameters.AddWithValue("@C", colName);
                        cmd.Parameters.AddWithValue("@Opt", optsStr);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("組合文字設定儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadMultiSelectConfigs();
                RefreshMultiConfiguredList();
            } catch (Exception ex) { MessageBox.Show("儲存失敗：" + ex.Message, "錯誤"); }
        }

        private void BtnClearMultiDb_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("【極度危險】\n您確定要永久刪除資料庫中「所有」的組合文字(複選)設定嗎？\n此操作無法復原！", "永久刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
            {
                string authPrompt = "清除設定資料需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (!AuthManager.VerifyAdmin(authPrompt)) return;

                try 
                {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                    {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM MultiSelectConfigs", conn)) 
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    
                    MessageBox.Show("資料庫內的所有組合文字設定已全部清空！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    _txtOptionsMulti.Clear();
                    LoadMultiSelectConfigs();
                    RefreshMultiConfiguredList();
                } 
                catch (Exception ex) 
                {
                    MessageBox.Show($"清除失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnExportMulti_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統組合文字(複選)設定_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                        {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT TableName AS [資料表名稱], ColName AS [欄位名稱], Options AS [選項內容(逗號分隔)] FROM MultiSelectConfigs", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        using (ExcelPackage p = new ExcelPackage()) 
                        {
                            var ws = p.Workbook.Worksheets.Add("組合文字設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("組合文字設定匯出成功！請直接在 Excel 中編輯後匯入。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnImportMulti_Click(object sender, EventArgs e)
        {
            string authPrompt = "匯入組合文字設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的設定檔" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    try 
                    {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) 
                        {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) 
                            {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) 
                                {
                                    for (int r = 2; r <= ws.Dimension.Rows; r++) 
                                    {
                                        string tb = ws.Cells[r, 1].Text.Trim();
                                        string col = ws.Cells[r, 2].Text.Trim();
                                        string opt = ws.Cells[r, 3].Text.Trim();

                                        if (string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(col)) continue;

                                        string sql = @"INSERT INTO MultiSelectConfigs (TableName, ColName, Options) 
                                                       VALUES (@T, @C, @Opt) 
                                                       ON CONFLICT(TableName, ColName) DO UPDATE SET Options=@Opt";
                                        
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) 
                                        {
                                            cmd.Parameters.AddWithValue("@T", tb);
                                            cmd.Parameters.AddWithValue("@C", col);
                                            cmd.Parameters.AddWithValue("@Opt", opt);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        
                        MessageBox.Show("組合文字設定已批次匯入並【自動存檔覆寫】成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                        _txtOptionsMulti.Clear();
                        LoadMultiSelectConfigs();
                        RefreshMultiConfiguredList();
                    } 
                    catch (Exception ex) 
                    {
                        MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnDelMulti_Click(object sender, EventArgs e)
        {
            var tb = _cboTableMulti.SelectedItem as ItemMap;
            if (tb == null || _cboColMulti.SelectedItem == null) return;

            if (MessageBox.Show("確定要刪除此欄位的組合文字設定嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                string authPrompt = "刪除組合文字設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                if (!AuthManager.VerifyAdmin(authPrompt)) return;
                
                try {
                    using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("DELETE FROM MultiSelectConfigs WHERE TableName=@T AND ColName=@C", conn)) {
                            cmd.Parameters.AddWithValue("@T", tb.EnName);
                            cmd.Parameters.AddWithValue("@C", _cboColMulti.SelectedItem.ToString());
                            cmd.ExecuteNonQuery();
                        }
                    }
                    _txtOptionsMulti.Clear();
                    LoadMultiSelectConfigs();
                    RefreshMultiConfiguredList();
                    MessageBox.Show("刪除成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message, "錯誤"); }
            }
        }

        private void RefreshMultiConfiguredList()
        {
            if (_flpMultiConfigured == null) return;
            _flpMultiConfigured.Controls.Clear();
            _selectedMultiItemPanel = null; 
            
            if (MultiSelectCache.Count == 0) {
                _flpMultiConfigured.Controls.Add(new Label { Text = "尚無任何設定。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) });
                return;
            }

            foreach (var kvp in MultiSelectCache) {
                string[] parts = kvp.Key.Split('|');
                if (parts.Length != 2) continue;

                string tbName = parts[0];
                string colName = parts[1];
                string chTbName = tbName;
                string dbName = ""; 

                foreach (var db in _dbMap) {
                    if (db.Value.Tables.ContainsKey(tbName)) {
                        chTbName = db.Value.Tables[tbName];
                        dbName = db.Key;
                        break;
                    }
                }

                Panel pItem = new Panel { Width = 650, Height = 45, BackColor = Color.AliceBlue, Margin = new Padding(5), Cursor = Cursors.Hand };
                
                Button btnDel = new Button { 
                    Text = "❌", 
                    Location = new Point(10, 7), 
                    Size = new Size(30, 30), 
                    FlatStyle = FlatStyle.Flat, 
                    ForeColor = Color.IndianRed, 
                    BackColor = Color.Transparent,
                    Cursor = Cursors.Hand,
                    Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold)
                };
                btnDel.FlatAppearance.BorderSize = 0;
                btnDel.FlatAppearance.MouseOverBackColor = Color.MistyRose;
                
                btnDel.Click += (s, e) => {
                    if (MessageBox.Show($"確定要刪除【{chTbName} - {colName}】的組合文字設定嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        string authPrompt = "刪除組合文字設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
                        if (!AuthManager.VerifyAdmin(authPrompt)) return;
                        
                        try {
                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var cmd = new SQLiteCommand("DELETE FROM MultiSelectConfigs WHERE TableName=@T AND ColName=@C", conn)) {
                                    cmd.Parameters.AddWithValue("@T", tbName);
                                    cmd.Parameters.AddWithValue("@C", colName);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            _txtOptionsMulti.Clear();
                            LoadMultiSelectConfigs();
                            RefreshMultiConfiguredList();
                            MessageBox.Show("刪除成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        } catch (Exception ex) { MessageBox.Show("刪除失敗：" + ex.Message, "錯誤"); }
                    }
                };

                Label lName = new Label { 
                    Text = $"表：{chTbName}   ➡️   欄位：[{colName}]", 
                    Location = new Point(50, 12), 
                    AutoSize = true, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                    ForeColor = Color.DarkSlateBlue,
                    Cursor = Cursors.Hand
                };

                Action selectAction = () => {
                    if (_selectedMultiItemPanel != null && _selectedMultiItemPanel != pItem) {
                        _selectedMultiItemPanel.BackColor = Color.AliceBlue; 
                    }
                    pItem.BackColor = Color.LightSkyBlue; 
                    _selectedMultiItemPanel = pItem;

                    foreach (ItemMap item in _cboDbMulti.Items) {
                        if (item.EnName == dbName) { _cboDbMulti.SelectedItem = item; break; }
                    }
                    foreach (ItemMap item in _cboTableMulti.Items) {
                        if (item.EnName == tbName) { _cboTableMulti.SelectedItem = item; break; }
                    }
                    if (_cboColMulti.Items.Contains(colName)) {
                        _cboColMulti.SelectedItem = colName;
                    }
                };

                pItem.Click += (s, e) => selectAction();
                lName.Click += (s, e) => selectAction();

                pItem.Controls.Add(btnDel);
                pItem.Controls.Add(lName);
                _flpMultiConfigured.Controls.Add(pItem);
            }
            
            _flpMultiConfigured.Resize -= FlpMultiConfigured_Resize; 
            _flpMultiConfigured.Resize += FlpMultiConfigured_Resize;
        }

        private void FlpMultiConfigured_Resize(object sender, EventArgs e)
        {
            foreach (Control ctrl in _flpMultiConfigured.Controls)
            {
                if (ctrl is Panel pnl)
                {
                    pnl.Width = _flpMultiConfigured.ClientSize.Width - 20;
                }
            }
        }

        public static void LoadMultiSelectConfigs()
        {
            MultiSelectCache.Clear();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmdChk = new SQLiteCommand("CREATE TABLE IF NOT EXISTS [MultiSelectConfigs] (Id INTEGER PRIMARY KEY AUTOINCREMENT, TableName TEXT, ColName TEXT, Options TEXT, UNIQUE(TableName, ColName));", conn)) { cmdChk.ExecuteNonQuery(); }

                    using (var cmd = new SQLiteCommand("SELECT TableName, ColName, Options FROM MultiSelectConfigs", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) {
                            string key = $"{reader["TableName"]}|{reader["ColName"]}";
                            string[] opts = reader["Options"].ToString().Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
                            MultiSelectCache[key] = opts;
                        }
                    }
                }
            } catch { }
        }

        public static string[] GetMultiSelectOptions(string tableName, string colName)
        {
            string key = $"{tableName}|{colName}";
            if (MultiSelectCache.ContainsKey(key)) return MultiSelectCache[key];
            return null;
        }
    }
}
