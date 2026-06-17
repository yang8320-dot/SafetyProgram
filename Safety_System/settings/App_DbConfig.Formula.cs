/// FILE: Safety_System/settings/App_DbConfig.Formula.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Safety_System
{
    public partial class App_DbConfig
    {
        private Label _lblFStartM, _lblFStartD, _lblFEndM, _lblFEndD;
        private ComboBox _cboFormulaType; 
        private NumericUpDown _numDecimals; 
        private ComboBox _cboRoundingMode;  
        private FlowLayoutPanel _pnlOps;  
        
        // 🟢 跨表變數生成器 UI 元件
        private Panel _pnlCrossTableBuilder;
        private ComboBox _cboCrossDb, _cboCrossTb, _cboCrossCol, _cboCrossAgg;
        
        private bool _isChangingFormulaDb = false; 

        private void BuildFormulaTab(TabPage tabFormula)
        {
            Panel pnlFormula = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

            GroupBox boxFormula = new GroupBox { Text = "資料表欄位自訂運算 (支援數學運算與文字組合)", Dock = DockStyle.Top, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            FlowLayoutPanel flpRow1 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 10) };
            Label lblFDb = new Label { Text = "選擇資料庫:", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaDb = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F), Margin = new Padding(0, 0, 30, 0) };
            _cboFormulaDb.SelectedIndexChanged += CboFormulaDb_SelectedIndexChanged;

            Label lblFTable = new Label { Text = "選擇資料表:", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaTable = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaTable.SelectedIndexChanged += CboFormulaTable_SelectedIndexChanged;

            flpRow1.Controls.AddRange(new Control[] { lblFDb, _cboFormulaDb, lblFTable, _cboFormulaTable });

            FlowLayoutPanel flpRow2 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 10) };
            Label lblFMatch = new Label { Text = "對應日期欄位：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaMatchCol = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblFStart = new Label { Text = "起：", AutoSize = true, Margin = new Padding(20, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFStartYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblFStartY = new Label { Text = "年", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFStartMonth = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFStartM = new Label { Text = "月", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFStartDay = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFStartD = new Label { Text = "日", AutoSize = true, Margin = new Padding(2, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblTilde = new Label { Text = "~", AutoSize = true, Margin = new Padding(5, 5, 10, 0), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold) };

            Label lblFEnd = new Label { Text = "迄：", AutoSize = true, Margin = new Padding(0, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFEndYear = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            Label lblFEndY = new Label { Text = "年", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFEndMonth = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFEndM = new Label { Text = "月", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFEndDay = new ComboBox { Width = 60, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _lblFEndD = new Label { Text = "日", AutoSize = true, Margin = new Padding(2, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };

            InitFormulaDateComboBoxes();

            _cboFormulaMatchCol.SelectedIndexChanged += (s, e) => {
                if (_cboFormulaMatchCol.SelectedItem != null) {
                    string sel = _cboFormulaMatchCol.SelectedItem.ToString();
                    bool showMonth = true; bool showDay = true;
                    if (sel == "年度" || sel.EndsWith("年")) { showMonth = false; showDay = false; } 
                    else if (sel == "年月" || sel == "月份" || sel.EndsWith("月")) { showMonth = true; showDay = false; }
                    _cboFStartMonth.Visible = showMonth; _lblFStartM.Visible = showMonth;
                    _cboFEndMonth.Visible = showMonth; _lblFEndM.Visible = showMonth;
                    _cboFStartDay.Visible = showDay; _lblFStartD.Visible = showDay;
                    _cboFEndDay.Visible = showDay; _lblFEndD.Visible = showDay;
                }
            };

            Button btnClearTime = new Button { Text = "♾️ 無限區間", Width = 110, Height = 32, Margin = new Padding(15, 0, 0, 0), BackColor = Color.LightGray, Cursor = Cursors.Hand };
            btnClearTime.Click += (s, e) => {
                _cboFStartYear.SelectedIndex = 0; _cboFStartMonth.SelectedIndex = 0; _cboFStartDay.SelectedIndex = 0;
                _cboFEndYear.SelectedIndex = _cboFEndYear.Items.Count - 1; _cboFEndMonth.SelectedIndex = 11; _cboFEndDay.SelectedIndex = _cboFEndDay.Items.Count - 1;
            };

            flpRow2.Controls.AddRange(new Control[] { 
                lblFMatch, _cboFormulaMatchCol, 
                lblFStart, _cboFStartYear, lblFStartY, _cboFStartMonth, _lblFStartM, _cboFStartDay, _lblFStartD,
                lblTilde,
                lblFEnd, _cboFEndYear, lblFEndY, _cboFEndMonth, _lblFEndM, _cboFEndDay, _lblFEndD,
                btnClearTime
            });

            FlowLayoutPanel flpRow3 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Padding = new Padding(0, 10, 0, 10) };
            Label lblFTarget = new Label { Text = "公式結果寫入至此欄：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaTargetCol = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblFType = new Label { Text = "公式類別：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboFormulaType = new ComboBox { Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };
            _cboFormulaType.Items.AddRange(new string[] { "數學運算", "組合文字", "運算+文字" });
            _cboFormulaType.SelectedIndex = 0;

            Label lblFDec = new Label { Text = "小數點：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _numDecimals = new NumericUpDown { Width = 50, Minimum = 0, Maximum = 4, Value = 4, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblFRound = new Label { Text = "進位：", AutoSize = true, Margin = new Padding(15, 5, 5, 0), Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboRoundingMode = new ComboBox { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboRoundingMode.Items.AddRange(new string[] { "四捨五入", "無條件進位", "無條件捨去" });
            _cboRoundingMode.SelectedIndex = 0;

            flpRow3.Controls.AddRange(new Control[] { lblFTarget, _cboFormulaTargetCol, lblFType, _cboFormulaType, lblFDec, _numDecimals, lblFRound, _cboRoundingMode });

            FlowLayoutPanel flpFormulaBlock = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(15, 10, 10, 15) };

            // 🟢 公式編輯器提示字眼動態變更
            Label lblFormula = new Label { 
                Text = "編輯公式：", 
                AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 10), ForeColor = Color.DarkCyan 
            };

            // 🟢 UI 連動邏輯：開啟對應的工具列與提示
            _cboFormulaType.SelectedIndexChanged += (s, e) => {
                string selType = _cboFormulaType.SelectedItem.ToString();
                bool showMathTools = (selType == "數學運算" || selType == "運算+文字");
                
                _pnlOps.Visible = showMathTools;
                _numDecimals.Visible = showMathTools;
                lblFDec.Visible = showMathTools;
                _cboRoundingMode.Visible = showMathTools;
                lblFRound.Visible = showMathTools;

                _pnlCrossTableBuilder.Visible = (selType == "運算+文字");

                if (selType == "運算+文字") {
                    lblFormula.Text = "編輯公式：(請將需要「數學運算」的欄位放於 { 大括號 } 內！)\n範例輸入： 廢回收率: { [回收水] / [總用水] * 100 } %";
                } else if (selType == "組合文字") {
                    lblFormula.Text = "編輯公式：(純粹將多個欄位與文字串接起來，不進行加減乘除)\n範例輸入： [年]-[月]-[日] 廠區代碼:[廠區]";
                } else {
                    lblFormula.Text = "編輯公式：(純數學運算，產出純數字)\n範例輸入： ([回收水] + [地下水]) / [總用水量]";
                }
            };

            FlowLayoutPanel pnlActionTools = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 0, 10) };
            
            _pnlOps = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            // 🟢 運算子加入大括號，方便使用者快速插入
            string[] ops = { "+", "-", "*", "/", "(", ")", "{", "}" };
            foreach (string op in ops) {
                Button b = new Button { Text = op, Width = 45, Height = 35, Font = new Font("Consolas", 14F, FontStyle.Bold), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke };
                b.Click += (s, e) => { _rtbFormulaEditor.Focus(); _rtbFormulaEditor.SelectedText = $" {op} "; };
                _pnlOps.Controls.Add(b);
            }

            Button btnInsertVar = new Button { Text = "插入本地欄位變數", Size = new Size(170, 35), BackColor = Color.LightSlateGray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnInsertVar.FlatAppearance.BorderSize = 0;
            btnInsertVar.Click += (s, e) => {
                using(Form fSel = new Form { Text = "選擇欄位", Size = new Size(300, 400), StartPosition = FormStartPosition.CenterParent }) {
                    ListBox lb = new ListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
                    if (_cboFormulaDb.SelectedItem != null && _cboFormulaTable.SelectedItem != null) {
                        string dName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
                        string tName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;
                        var cols = DataManager.GetColumnNames(dName, tName);
                        foreach(var c in cols) lb.Items.Add(c);
                    }
                    lb.DoubleClick += (s2, e2) => {
                        if (lb.SelectedItem != null) { _rtbFormulaEditor.Focus(); _rtbFormulaEditor.SelectedText = $"[{lb.SelectedItem}]"; fSel.Close(); }
                    };
                    fSel.Controls.Add(lb); fSel.ShowDialog();
                }
            };
            
            Button btnClearForm = new Button { Text = "✨ 清空編輯區重新設定", Size = new Size(200, 35), BackColor = Color.Gainsboro, ForeColor = Color.Black, FlatStyle = FlatStyle.Flat, Margin = new Padding(10, 0, 0, 0) };
            btnClearForm.Click += (s, e) => {
                _currentFormulaEditId = 0; _rtbFormulaEditor.Clear(); _cboFormulaTargetCol.SelectedIndex = -1; _cboFormulaMatchCol.SelectedIndex = -1;
                _cboFormulaType.SelectedIndex = 0;
            };

            pnlActionTools.Controls.Add(_pnlOps);
            pnlActionTools.Controls.Add(btnInsertVar);
            pnlActionTools.Controls.Add(btnClearForm);

            _pnlCrossTableBuilder = new Panel { Width = 950, Height = 45, BackColor = Color.FromArgb(240, 245, 250), Margin = new Padding(0, 0, 0, 10), Visible = false };
            _pnlCrossTableBuilder.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, _pnlCrossTableBuilder.ClientRectangle, Color.LightSteelBlue, ButtonBorderStyle.Solid);
            
            Label lblCross = new Label { Text = "插入跨表運算：", AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue };
            _cboCrossDb = new ComboBox { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(135, 10) };
            _cboCrossTb = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(275, 10) };
            _cboCrossCol = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(465, 10) };
            
            _cboCrossAgg = new ComboBox { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 11F), Location = new Point(655, 10) };
            _cboCrossAgg.Items.AddRange(new string[] { "SUM", "AVG", "MAX", "MIN", "COUNT" });
            _cboCrossAgg.SelectedIndex = 0;

            Button btnInsertCross = new Button { Text = "插入 ⬇️", Width = 90, Height = 32, Location = new Point(765, 8), BackColor = Color.DarkCyan, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnInsertCross.FlatAppearance.BorderSize = 0;

            _pnlCrossTableBuilder.Controls.AddRange(new Control[] { lblCross, _cboCrossDb, _cboCrossTb, _cboCrossCol, _cboCrossAgg, btnInsertCross });

            _cboCrossDb.Items.Add(new ItemMap { EnName = "", ChName = "" });
            var dbMap = App_DbConfig.GetDbMapCache();
            foreach (var kvp in dbMap) _cboCrossDb.Items.Add(new ItemMap { EnName = kvp.Key, ChName = kvp.Value.ChDbName });

            _cboCrossDb.SelectedIndexChanged += (s, e) => {
                _cboCrossTb.Items.Clear(); _cboCrossTb.Items.Add(new ItemMap { EnName = "", ChName = "" });
                var db = _cboCrossDb.SelectedItem as ItemMap;
                if (db != null && !string.IsNullOrEmpty(db.EnName)) {
                    foreach (var tb in dbMap[db.EnName].Tables) _cboCrossTb.Items.Add(new ItemMap { EnName = tb.Key, ChName = tb.Value });
                }
            };

            _cboCrossTb.SelectedIndexChanged += (s, e) => {
                _cboCrossCol.Items.Clear();
                var db = _cboCrossDb.SelectedItem as ItemMap; var tb = _cboCrossTb.SelectedItem as ItemMap;
                if (db != null && tb != null && !string.IsNullOrEmpty(db.EnName) && !string.IsNullOrEmpty(tb.EnName)) {
                    var cols = DataManager.GetColumnNames(db.EnName, tb.EnName).Where(c => c != "Id");
                    foreach (var c in cols) _cboCrossCol.Items.Add(c);
                }
            };

            btnInsertCross.Click += (s, e) => {
                var db = _cboCrossDb.SelectedItem as ItemMap; var tb = _cboCrossTb.SelectedItem as ItemMap;
                if (db == null || tb == null || _cboCrossCol.SelectedItem == null) { MessageBox.Show("請選擇完整的跨表欄位來源！"); return; }
                _rtbFormulaEditor.Focus();
                _rtbFormulaEditor.SelectedText = $"{_cboCrossAgg.Text}([{db.EnName}].[{tb.EnName}].[{_cboCrossCol.Text}])";
            };

            _rtbFormulaEditor = new RichTextBox { Width = 950, Height = 120, Font = new Font("Consolas", 14F), BackColor = Color.AliceBlue, Margin = new Padding(0, 0, 0, 15) };

            Button btnSaveFormula = new Button { Text = "💾 儲存此運算公式", Size = new Size(200, 45), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnSaveFormula.FlatAppearance.BorderSize = 0;
            btnSaveFormula.Click += BtnSaveFormula_Click;

            flpFormulaBlock.Controls.AddRange(new Control[] { lblFormula, pnlActionTools, _pnlCrossTableBuilder, _rtbFormulaEditor, btnSaveFormula });

            boxFormula.Controls.Add(flpFormulaBlock);
            boxFormula.Controls.Add(flpRow3);
            boxFormula.Controls.Add(flpRow2);
            boxFormula.Controls.Add(flpRow1);

            TableLayoutPanel tlpListArea = new TableLayoutPanel { Dock = DockStyle.Top, Height = 400, ColumnCount = 1, RowCount = 2, Margin = new Padding(0, 20, 0, 0) };
            tlpListArea.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlpListArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            FlowLayoutPanel pnlFormulaAction = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(0, 0, 0, 5), FlowDirection = FlowDirection.LeftToRight };

            Button btnExportFormula = new Button { Text = "📤 匯出所有公式", Width = 160, Height = 40, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 15, 0) };
            btnExportFormula.FlatAppearance.BorderSize = 0;
            btnExportFormula.Click += BtnExportFormula_Click;

            Button btnImportFormula = new Button { Text = "📥 匯入公式設定", Width = 160, Height = 40, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 15, 0) };
            btnImportFormula.FlatAppearance.BorderSize = 0;
            btnImportFormula.Click += BtnImportFormula_Click;

            Button btnRecalculateAll = new Button { Text = "🔄 全面更新歷史數據", Width = 200, Height = 40, BackColor = Color.DarkOrange, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Margin = new Padding(0) };
            btnRecalculateAll.FlatAppearance.BorderSize = 0;
            btnRecalculateAll.Click += async (s, e) => {
                if (_cboFormulaDb.SelectedItem == null || _cboFormulaTable.SelectedItem == null) {
                    MessageBox.Show("請在最上方「選擇資料庫」與「選擇資料表」，系統才知道要重算哪一張表喔！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                }
                string dbName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
                string tableName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;
                if (string.IsNullOrEmpty(dbName) || string.IsNullOrEmpty(tableName)) return;

                if (MessageBox.Show($"您確定要全面重新計算【{tableName}】的所有歷史資料嗎？\n\n系統將會套用下方清單中屬於該表的所有公式，並將結果回寫至資料庫。", "確認全面更新", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    await RunBackgroundRecalculation(dbName, tableName);
                }
            };

            pnlFormulaAction.Controls.Add(btnExportFormula);
            pnlFormulaAction.Controls.Add(btnImportFormula);
            pnlFormulaAction.Controls.Add(btnRecalculateAll); 

            GroupBox boxFormulasList = new GroupBox { Text = "該表目前的公式清單 (點擊可編輯)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            _flpFormulasList = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxFormulasList.Controls.Add(_flpFormulasList);

            tlpListArea.Controls.Add(pnlFormulaAction, 0, 0);
            tlpListArea.Controls.Add(boxFormulasList, 0, 1);

            pnlFormula.Controls.Add(tlpListArea);
            pnlFormula.Controls.Add(boxFormula);
            tabFormula.Controls.Add(pnlFormula);

            tabFormula.Enter += (s, e) => { RefreshAllFormulasList(); };
        }

        private void InitFormulaDateComboBoxes()
        {
            _cboFStartYear.Items.Add("1900"); _cboFEndYear.Items.Add("1900");
            int currY = DateTime.Today.Year;
            for (int i = currY - 10; i <= currY + 1; i++) {
                _cboFStartYear.Items.Add(i.ToString()); _cboFEndYear.Items.Add(i.ToString());
            }
            _cboFStartYear.Items.Add("2099"); _cboFEndYear.Items.Add("2099");

            for (int i = 1; i <= 12; i++) {
                string m = i.ToString("D2");
                _cboFStartMonth.Items.Add(m); _cboFEndMonth.Items.Add(m);
            }

            _cboFStartYear.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFStartYear, _cboFStartMonth, _cboFStartDay);
            _cboFStartMonth.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFStartYear, _cboFStartMonth, _cboFStartDay);
            _cboFEndYear.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFEndYear, _cboFEndMonth, _cboFEndDay);
            _cboFEndMonth.SelectedIndexChanged += (s, e) => UpdateFormulaDaysCombo(_cboFEndYear, _cboFEndMonth, _cboFEndDay);

            SetFormulaDateStr(DateTime.Today.ToString("yyyy-01-01"), _cboFStartYear, _cboFStartMonth, _cboFStartDay);
            SetFormulaDateStr(DateTime.Today.ToString("yyyy-12-31"), _cboFEndYear, _cboFEndMonth, _cboFEndDay);
        }

        private void UpdateFormulaDaysCombo(ComboBox y, ComboBox m, ComboBox d)
        {
            if (y.SelectedItem == null || m.SelectedItem == null) return;
            if (!int.TryParse(y.SelectedItem.ToString(), out int year) || !int.TryParse(m.SelectedItem.ToString(), out int month)) return;
            int days = DateTime.DaysInMonth(year, month);
            string currentDay = d.SelectedItem?.ToString();
            d.Items.Clear();
            for (int i = 1; i <= days; i++) d.Items.Add(i.ToString("D2"));
            if (currentDay != null && d.Items.Contains(currentDay)) d.SelectedItem = currentDay;
            else d.SelectedIndex = d.Items.Count - 1;
        }

        private string GetFormulaDateStr(ComboBox y, ComboBox m, ComboBox d, string matchCol)
        {
            string year = y.SelectedItem?.ToString() ?? "1900";
            string month = m.SelectedItem?.ToString() ?? "01";
            string day = d.SelectedItem?.ToString() ?? "01";
            if (string.IsNullOrEmpty(matchCol)) return $"{year}-{month}-{day}";
            if (matchCol == "年度" || matchCol.EndsWith("年")) return year;
            if (matchCol == "年月" || matchCol == "月份" || matchCol.EndsWith("月")) return $"{year}-{month}";
            return $"{year}-{month}-{day}";
        }

        private void SetFormulaDateStr(string dateStr, ComboBox y, ComboBox m, ComboBox d)
        {
            if (string.IsNullOrEmpty(dateStr)) return;
            var parts = dateStr.Split('-');
            if (parts.Length >= 1 && y.Items.Contains(parts[0])) y.SelectedItem = parts[0];
            if (parts.Length >= 2 && m.Items.Contains(parts[1])) m.SelectedItem = parts[1];
            if (parts.Length >= 3) {
                UpdateFormulaDaysCombo(y, m, d);
                if (d.Items.Contains(parts[2])) d.SelectedItem = parts[2];
            }
        }

        private void CboFormulaDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isChangingFormulaDb) return; 

            _cboFormulaTable.Items.Clear();
            _cboFormulaTable.Items.Add(new ItemMap { EnName = "", ChName = "" });
            if (_cboFormulaDb.SelectedItem == null) return;

            var selectedDb = (ItemMap)_cboFormulaDb.SelectedItem;

            if (selectedDb.EnName.StartsWith("Menu") && selectedDb.EnName.EndsWith("DB")) {
                string menuName = selectedDb.EnName.Replace("DB", "").Replace("Menu", "選單");
                if (!AuthManager.VerifyHiddenMenu(menuName)) {
                    _isChangingFormulaDb = true;
                    _cboFormulaDb.SelectedIndex = 0; 
                    _isChangingFormulaDb = false;
                    return;
                }
            }

            if (!string.IsNullOrEmpty(selectedDb.EnName) && _dbMap.ContainsKey(selectedDb.EnName)) {
                var tbItems = _dbMap[selectedDb.EnName].Tables.Select(tbl => new ItemMap { EnName = tbl.Key, ChName = tbl.Value }).ToArray();
                _cboFormulaTable.Items.AddRange(tbItems);
            }
        }

        private void CboFormulaTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            _cboFormulaTargetCol.Items.Clear();
            _cboFormulaMatchCol.Items.Clear();
            _cboFormulaMatchCol.Items.Add(""); 
            if (_cboFormulaDb.SelectedItem == null || _cboFormulaTable.SelectedItem == null) return;
            string dbName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;

            if (!string.IsNullOrEmpty(dbName) && !string.IsNullOrEmpty(tableName)) {
                var cols = DataManager.GetColumnNames(dbName, tableName).Where(c => c != "Id").ToArray();
                _cboFormulaTargetCol.Items.AddRange(cols);
                _cboFormulaMatchCol.Items.AddRange(cols);
            }

            RefreshAllFormulasList();
        }

        private async void BtnSaveFormula_Click(object sender, EventArgs e)
        {
            if (_cboFormulaDb.SelectedItem == null || _cboFormulaTable.SelectedItem == null || _cboFormulaTargetCol.SelectedItem == null) {
                MessageBox.Show("請確認資料庫、資料表與目標欄位皆已選擇！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string dbName = ((ItemMap)_cboFormulaDb.SelectedItem).EnName;
            string tableName = ((ItemMap)_cboFormulaTable.SelectedItem).EnName;
            string targetCol = _cboFormulaTargetCol.SelectedItem.ToString();
            string matchCol = _cboFormulaMatchCol.SelectedItem?.ToString() ?? "";
            
            string sDate = GetFormulaDateStr(_cboFStartYear, _cboFStartMonth, _cboFStartDay, matchCol);
            string eDate = GetFormulaDateStr(_cboFEndYear, _cboFEndMonth, _cboFEndDay, matchCol);
            
            string formulaType = _cboFormulaType.SelectedItem.ToString();
            string formula = _rtbFormulaEditor.Text.Trim();
            
            int decPlaces = (int)_numDecimals.Value;
            string roundMode = _cboRoundingMode.SelectedItem.ToString();

            if (string.IsNullOrEmpty(formula)) {
                MessageBox.Show("公式不可為空！若要取消該欄位的公式，請在下方清單點擊刪除。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            if (string.Compare(sDate, eDate) > 0) {
                MessageBox.Show("【起日】不能大於【迄日】！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string checkSql = "SELECT StartDate, EndDate FROM ColumnFormulas WHERE DbName=@DB AND TableName=@TB AND TargetCol=@TC AND Id != @ExId";
                    using (var cmd = new SQLiteCommand(checkSql, conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                        cmd.Parameters.AddWithValue("@TC", targetCol); cmd.Parameters.AddWithValue("@ExId", _currentFormulaEditId); 
                        using (var reader = cmd.ExecuteReader()) {
                            while (reader.Read()) {
                                string os = reader["StartDate"].ToString(); string oe = reader["EndDate"].ToString();
                                if (string.Compare(sDate, oe) <= 0 && string.Compare(eDate, os) >= 0) {
                                    MessageBox.Show($"此目標欄位在該時間區間內已有設定其他公式！\n重疊的區間：{os} ~ {oe}\n\n為避免資料計算衝突，請強制調整您的時間區間。", "重疊防呆攔截", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }

                    if (_currentFormulaEditId > 0) {
                        string updateSql = "UPDATE ColumnFormulas SET MatchCol=@MC, StartDate=@SD, EndDate=@ED, FormulaType=@FT, Formula=@F, DecimalPlaces=@DP, RoundingMode=@RM WHERE Id=@Id";
                        using(var cmd = new SQLiteCommand(updateSql, conn)) {
                            cmd.Parameters.AddWithValue("@MC", matchCol); cmd.Parameters.AddWithValue("@SD", sDate);
                            cmd.Parameters.AddWithValue("@ED", eDate); cmd.Parameters.AddWithValue("@FT", formulaType);
                            cmd.Parameters.AddWithValue("@F", formula); cmd.Parameters.AddWithValue("@Id", _currentFormulaEditId);
                            cmd.Parameters.AddWithValue("@DP", decPlaces); cmd.Parameters.AddWithValue("@RM", roundMode);
                            cmd.ExecuteNonQuery();
                        }
                    } else {
                        string insertSql = "INSERT INTO ColumnFormulas (DbName, TableName, TargetCol, MatchCol, StartDate, EndDate, FormulaType, Formula, DecimalPlaces, RoundingMode) VALUES (@DB, @TB, @TC, @MC, @SD, @ED, @FT, @F, @DP, @RM)";
                        using(var cmd = new SQLiteCommand(insertSql, conn)) {
                            cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                            cmd.Parameters.AddWithValue("@TC", targetCol); cmd.Parameters.AddWithValue("@MC", matchCol);
                            cmd.Parameters.AddWithValue("@SD", sDate); cmd.Parameters.AddWithValue("@ED", eDate);
                            cmd.Parameters.AddWithValue("@FT", formulaType); cmd.Parameters.AddWithValue("@F", formula);
                            cmd.Parameters.AddWithValue("@DP", decPlaces); cmd.Parameters.AddWithValue("@RM", roundMode);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            } catch (Exception ex) { MessageBox.Show("儲存公式時發生錯誤：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            MessageBox.Show($"【{targetCol}】 運算公式已成功儲存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _currentFormulaEditId = 0; _rtbFormulaEditor.Clear(); RefreshAllFormulasList();

            if (MessageBox.Show($"公式已儲存。\n\n是否要立即在背景重新計算【{tableName}】的所有歷史資料？\n\n(系統將以低記憶體消耗方式逐筆刷新，並自動觸發相關的資料同步)", "背景重算確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                await RunBackgroundRecalculation(dbName, tableName);
            }
        }

        private void RefreshAllFormulasList()
        {
            if (_flpFormulasList == null) return;
            _flpFormulasList.Controls.Clear();
            
            string currentDb = _cboFormulaDb.SelectedItem != null ? ((ItemMap)_cboFormulaDb.SelectedItem).EnName : "";
            string currentTb = _cboFormulaTable.SelectedItem != null ? ((ItemMap)_cboFormulaTable.SelectedItem).EnName : "";
            
            if (string.IsNullOrEmpty(currentDb) || string.IsNullOrEmpty(currentTb)) {
                _flpFormulasList.Controls.Add(new Label { Text = "請先在上方選擇庫與表以檢視公式。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                return;
            }

            DataTable dt = new DataTable();
            try {
                using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM ColumnFormulas WHERE DbName=@DB AND TableName=@TB", conn)) {
                        cmd.Parameters.AddWithValue("@DB", currentDb);
                        cmd.Parameters.AddWithValue("@TB", currentTb);
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                    }
                }
            } catch { }

            if (dt.Rows.Count == 0) {
                _flpFormulasList.Controls.Add(new Label { Text = "該表目前沒有設定任何自動運算公式。", ForeColor = Color.DimGray, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) });
                return;
            }

            string chDbName = _cboFormulaDb.Text;
            string chTbName = _cboFormulaTable.Text;

            foreach (DataRow row in dt.Rows) {
                int id = Convert.ToInt32(row["Id"]);
                string targetCol = row["TargetCol"].ToString(); 
                string matchCol = row["MatchCol"].ToString();
                string sDate = row["StartDate"].ToString(); 
                string eDate = row["EndDate"].ToString();
                string fType = row.Table.Columns.Contains("FormulaType") ? row["FormulaType"].ToString() : "數學運算";
                string formula = row["Formula"].ToString();
                
                int decPlaces = row.Table.Columns.Contains("DecimalPlaces") && row["DecimalPlaces"] != DBNull.Value ? Convert.ToInt32(row["DecimalPlaces"]) : 4;
                string rMode = row.Table.Columns.Contains("RoundingMode") && row["RoundingMode"] != DBNull.Value ? row["RoundingMode"].ToString() : "四捨五入";

                string dateInfo = string.IsNullOrEmpty(matchCol) ? "" : $" (當 [{matchCol}] 介於 {sDate} ~ {eDate} 時)";
                string roundInfo = (fType == "數學運算" || fType == "運算+文字") ? $" (小數點:{decPlaces}位 | {rMode})" : "";

                string text = $"【{fType}】表:[{chTbName}]  ➡️  目標:[{targetCol}] = {formula}{dateInfo}{roundInfo}";
                
                Color fColor = Color.DarkSlateBlue;
                if (fType == "組合文字") fColor = Color.DarkCyan;
                else if (fType == "運算+文字") fColor = Color.DarkOrange;

                Label lTxt = new Label { Text = text, AutoSize = true, Location = new Point(10, 12), Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold), ForeColor = fColor, Cursor = Cursors.Hand };
                int reqW = TextRenderer.MeasureText(text, lTxt.Font).Width + 100;
                int panelW = Math.Max(_flpFormulasList.ClientSize.Width - 25, reqW);

                Panel p = new Panel { Width = panelW, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(5), Cursor = Cursors.Hand };
                Button btnDel = new Button { Text = "❌", Width = 40, Height = 35, Location = new Point(panelW - 60, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                btnDel.FlatAppearance.BorderSize = 0;
                
                btnDel.Click += (s, ev) => {
                    if (MessageBox.Show($"確定刪除此自動運算公式？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("DELETE FROM ColumnFormulas WHERE Id=@Id", conn)) {
                                cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery();
                            }
                        }
                        RefreshAllFormulasList();
                        _currentFormulaEditId = 0;
                    }
                };

                Action loadToEdit = () => {
                    _currentFormulaEditId = id;
                    if (_cboFormulaTargetCol.Items.Contains(targetCol)) _cboFormulaTargetCol.SelectedItem = targetCol;
                    if (_cboFormulaMatchCol.Items.Contains(matchCol)) _cboFormulaMatchCol.SelectedItem = matchCol;
                    SetFormulaDateStr(sDate, _cboFStartYear, _cboFStartMonth, _cboFStartDay);
                    SetFormulaDateStr(eDate, _cboFEndYear, _cboFEndMonth, _cboFEndDay);
                    
                    if (_cboFormulaType.Items.Contains(fType)) _cboFormulaType.SelectedItem = fType;
                    
                    _numDecimals.Value = decPlaces;
                    if (_cboRoundingMode.Items.Contains(rMode)) _cboRoundingMode.SelectedItem = rMode;
                    
                    _rtbFormulaEditor.Text = formula;
                };

                p.Click += (s, ev) => loadToEdit();
                lTxt.Click += (s, ev) => loadToEdit();

                p.Controls.Add(lTxt);
                p.Controls.Add(btnDel);
                _flpFormulasList.Controls.Add(p);
            }
        }

        private void BtnExportFormula_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx", FileName = "系統自動運算公式設定_" + DateTime.Now.ToString("yyyyMMdd") }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT DbName AS [資料庫名], TableName AS [資料表名], TargetCol AS [目標欄位], MatchCol AS [對應日期欄位], StartDate AS [區間起日], EndDate AS [區間迄日], FormulaType AS [公式類型], Formula AS [運算公式], DecimalPlaces AS [小數點], RoundingMode AS [進位模式] FROM ColumnFormulas", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        using (ExcelPackage p = new ExcelPackage()) {
                            var ws = p.Workbook.Worksheets.Add("自動運算公式設定");
                            ws.Cells["A1"].LoadFromDataTable(dt, true);
                            ws.Cells.AutoFitColumns();
                            p.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("自動運算公式設定匯出成功！您可以以此檔案作為備份。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) { MessageBox.Show("匯出失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnImportFormula_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "選擇要匯入的公式設定檔" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(ofd.FileName))) {
                            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault();
                            if (ws == null || ws.Dimension == null) return;

                            using (var conn = new SQLiteConnection($"Data Source={DataManager.SysConfigDbPath};Version=3;")) {
                                conn.Open();
                                using (var trans = conn.BeginTransaction()) {
                                    new SQLiteCommand("DELETE FROM ColumnFormulas", conn, trans).ExecuteNonQuery();

                                    for (int r = 2; r <= ws.Dimension.Rows; r++) {
                                        string db = ws.Cells[r, 1].Text.Trim(); string tb = ws.Cells[r, 2].Text.Trim();
                                        string targetCol = ws.Cells[r, 3].Text.Trim(); string matchCol = ws.Cells[r, 4].Text.Trim();
                                        string sDate = ws.Cells[r, 5].Text.Trim(); string eDate = ws.Cells[r, 6].Text.Trim();
                                        string fType = ws.Cells[r, 7].Text.Trim(); string formula = ws.Cells[r, 8].Text.Trim();
                                        string decStr = ws.Cells[r, 9].Text.Trim(); string rMode = ws.Cells[r, 10].Text.Trim();

                                        if (string.IsNullOrEmpty(db) || string.IsNullOrEmpty(tb) || string.IsNullOrEmpty(targetCol) || string.IsNullOrEmpty(formula)) continue;

                                        int decPlaces = string.IsNullOrEmpty(decStr) ? 4 : int.Parse(decStr);
                                        rMode = string.IsNullOrEmpty(rMode) ? "四捨五入" : rMode;

                                        string sql = @"INSERT INTO ColumnFormulas (DbName, TableName, TargetCol, MatchCol, StartDate, EndDate, FormulaType, Formula, DecimalPlaces, RoundingMode) VALUES (@DB, @TB, @TC, @MC, @SD, @ED, @FT, @F, @DP, @RM)";
                                        using (var cmd = new SQLiteCommand(sql, conn, trans)) {
                                            cmd.Parameters.AddWithValue("@DB", db); cmd.Parameters.AddWithValue("@TB", tb);
                                            cmd.Parameters.AddWithValue("@TC", targetCol); cmd.Parameters.AddWithValue("@MC", matchCol);
                                            cmd.Parameters.AddWithValue("@SD", sDate); cmd.Parameters.AddWithValue("@ED", eDate);
                                            cmd.Parameters.AddWithValue("@FT", string.IsNullOrEmpty(fType) ? "數學運算" : fType);
                                            cmd.Parameters.AddWithValue("@F", formula); 
                                            cmd.Parameters.AddWithValue("@DP", decPlaces); cmd.Parameters.AddWithValue("@RM", rMode);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                        }
                        MessageBox.Show("自動運算公式設定已批次匯入並覆寫成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        RefreshAllFormulasList();
                    } catch (Exception ex) { MessageBox.Show("匯入失敗，請確認檔案格式是否正確：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }
    }
}
