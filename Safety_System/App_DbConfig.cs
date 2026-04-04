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

        // 預先對應各資料庫的所有資料表，方便下拉選單切換
        private readonly Dictionary<string, string[]> _dbMap = new Dictionary<string, string[]> {
            { "Safety", new[] { "NearMiss", "SafetyInspection", "SafetyObservation", "TrafficInjury", "WorkInjury" } },
            { "Nursing", new[] { "HealthPromotion", "WorkInjuryReport" } },
            { "Air", new[] { "AirPollution" } },
            { "Water", new[] { "DischargeData", "WaterMeterReadings", "WaterChemicals", "WaterVolume" } },
            { "Waste", new[] { "WasteMonthly" } },
            { "Fire", new[] { "FireResponsible", "HazardStats", "FireEquip" } }
        };

        public Control GetView()
        {
            Panel main = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

            // 🟢 第一區塊：資料庫存放路徑設定 (框內與文字間隔 15)
            GroupBox boxPath = new GroupBox { Text = "資料庫存放路徑設定", Dock = DockStyle.Top, Height = 140, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
            _txtPath = new TextBox { Location = new Point(20, 40), Width = 450, ReadOnly = true, Text = DataManager.BasePath, Font = new Font("Microsoft JhengHei UI", 12F) };
            Button btnBrowse = new Button { Text = "選擇資料夾", Location = new Point(480, 38), Size = new Size(120, 35), Font = new Font("Microsoft JhengHei UI", 12F) };
            btnBrowse.Click += BtnBrowse_Click;
            
            Button btnSavePath = new Button { Text = "儲存路徑變更", Location = new Point(20, 85), Size = new Size(180, 40), BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSavePath.Click += BtnSavePath_Click;

            boxPath.Controls.AddRange(new Control[] { _txtPath, btnBrowse, btnSavePath });

            // 🟢 第二區塊：防重寫欄位設定 (複合鍵)
            GroupBox boxKeys = new GroupBox { Text = "資料表防重寫欄位設定 (空值則正常寫入不防呆)", Dock = DockStyle.Top, Height = 250, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            // 利用 Margin 推開與上方 GroupBox 的距離
            boxKeys.Margin = new Padding(0, 20, 0, 0);

            Label lblDb = new Label { Text = "選擇資料庫:", Location = new Point(20, 40), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboDb = new ComboBox { Location = new Point(130, 38), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            foreach (var key in _dbMap.Keys) _cboDb.Items.Add(key);
            _cboDb.SelectedIndexChanged += CboDb_SelectedIndexChanged;

            Label lblTable = new Label { Text = "選擇資料表:", Location = new Point(330, 40), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable = new ComboBox { Location = new Point(440, 38), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboTable.SelectedIndexChanged += CboTable_SelectedIndexChanged;

            Label lblCol1 = new Label { Text = "判斷欄位一:", Location = new Point(20, 90), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol1 = new ComboBox { Location = new Point(130, 88), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCol2 = new Label { Text = "判斷欄位二:", Location = new Point(330, 90), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
            _cboCol2 = new ComboBox { Location = new Point(440, 88), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            Button btnSaveKeys = new Button { Text = "儲存防重寫規則", Location = new Point(20, 140), Size = new Size(180, 40), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveKeys.Click += BtnSaveKeys_Click;

            boxKeys.Controls.AddRange(new Control[] { lblDb, _cboDb, lblTable, _cboTable, lblCol1, _cboCol1, lblCol2, _cboCol2, btnSaveKeys });

            // 為了讓兩個 GroupBox 之間有間距，插入一個空白 Panel
            Panel spacer = new Panel { Dock = DockStyle.Top, Height = 20 };

            main.Controls.Add(boxKeys);
            main.Controls.Add(spacer);
            main.Controls.Add(boxPath);

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
            _cboTable.Items.Clear();
            _cboCol1.Items.Clear(); _cboCol2.Items.Clear();
            if (_cboDb.SelectedItem == null) return;
            string db = _cboDb.SelectedItem.ToString();
            if (_dbMap.ContainsKey(db)) {
                _cboTable.Items.AddRange(_dbMap[db]);
                if (_cboTable.Items.Count > 0) _cboTable.SelectedIndex = 0;
            }
        }

        private void CboTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            _cboCol1.Items.Clear(); _cboCol2.Items.Clear();
            _cboCol1.Items.Add(""); _cboCol2.Items.Add(""); // 加入空選項

            if (_cboDb.SelectedItem == null || _cboTable.SelectedItem == null) return;
            
            string dbName = _cboDb.SelectedItem.ToString();
            string tableName = _cboTable.SelectedItem.ToString();

            // 動態去資料庫抓欄位名稱
            List<string> cols = DataManager.GetColumnNames(dbName, tableName);
            foreach(var c in cols) {
                if (c != "Id") { _cboCol1.Items.Add(c); _cboCol2.Items.Add(c); }
            }

            // 載入目前設定的記憶值
            var keys = DataManager.GetTableKeys(dbName, tableName);
            if (!string.IsNullOrEmpty(keys.col1) && _cboCol1.Items.Contains(keys.col1)) _cboCol1.SelectedItem = keys.col1;
            else _cboCol1.SelectedIndex = 0;

            if (!string.IsNullOrEmpty(keys.col2) && _cboCol2.Items.Contains(keys.col2)) _cboCol2.SelectedItem = keys.col2;
            else _cboCol2.SelectedIndex = 0;
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
