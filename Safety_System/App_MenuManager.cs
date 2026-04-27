/// FILE: Safety_System/App_MenuManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_MenuManager : Form
    {
        private ComboBox _cboCategory;
        private TextBox _txtMenuName;
        private Button _btnAdd;
        private Button _btnRename;
        
        private FlowLayoutPanel _flpCustomMenus;

        // 定義可被擴充的主選單分類，及其對應的實體資料庫名稱
        private readonly Dictionary<string, string> _categoryToDbMap = new Dictionary<string, string>
        {
            { "日常作業", "Reports" },
            { "工安", "Safety" },
            { "化學品", "Chemical" },
            { "化學品要求及規範", "Chemical" },
            { "護理", "Nursing" },
            { "空污", "Air" },
            { "水污", "Water" },
            { "產能及廢棄物", "Waste" },
            { "消防", "Fire" },
            { "檢測數據", "TestData" },
            { "教育訓練", "教育訓練" },
            { "ESG", "ESG" },
            { "ISO14001", "ISO14001" },
            { "選單1", "Menu1DB" },
            { "選單2", "Menu2DB" },
            { "選單3", "Menu3DB" },
            { "選單4", "Menu4DB" }
        };

        public App_MenuManager()
        {
            // 初始化自訂選單對照表 (存放在系統資料庫)
            string createSql = "CREATE TABLE IF NOT EXISTS [CustomMenus] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [分類] TEXT, [資料庫名] TEXT, [資料表名] TEXT);";
            DataManager.InitTable("SystemConfig", "CustomMenus", createSql);
            
            InitializeComponent();
            RefreshCustomMenusList();
        }

        private void InitializeComponent()
        {
            this.Text = "選單新增與管理";
            this.Size = new Size(600, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label { Text = "📂 自訂選單新增與管理", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.SteelBlue, Location = new Point(30, 20), AutoSize = true };

            GroupBox boxAdd = new GroupBox { Text = "操作區 (新增 / 更名)", Location = new Point(30, 70), Size = new Size(520, 160), Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCat = new Label { Text = "目標分類：", Location = new Point(20, 40), AutoSize = true };
            _cboCategory = new ComboBox { Location = new Point(120, 37), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboCategory.Items.AddRange(_categoryToDbMap.Keys.ToArray());
            if (_cboCategory.Items.Count > 0) _cboCategory.SelectedIndex = 0;

            Label lblName = new Label { Text = "選單名稱：", Location = new Point(20, 95), AutoSize = true };
            _txtMenuName = new TextBox { Location = new Point(120, 92), Width = 180 };

            _btnAdd = new Button { Text = "➕ 新增", Location = new Point(320, 35), Size = new Size(80, 85), BackColor = Color.ForestGreen, ForeColor = Color.White, Cursor = Cursors.Hand };
            _btnAdd.Click += BtnAdd_Click;

            _btnRename = new Button { Text = "✏️ 更名", Location = new Point(410, 35), Size = new Size(80, 85), BackColor = Color.DarkOrange, ForeColor = Color.White, Cursor = Cursors.Hand };
            _btnRename.Click += BtnRename_Click;

            boxAdd.Controls.Add(lblCat);
            boxAdd.Controls.Add(_cboCategory);
            boxAdd.Controls.Add(lblName);
            boxAdd.Controls.Add(_txtMenuName);
            boxAdd.Controls.Add(_btnAdd);
            boxAdd.Controls.Add(_btnRename);

            GroupBox boxList = new GroupBox { Text = "已建立之自訂選單 (刪除)", Location = new Point(30, 250), Size = new Size(520, 390), Font = new Font("Microsoft JhengHei UI", 12F) };
            
            _flpCustomMenus = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            boxList.Controls.Add(_flpCustomMenus);

            this.Controls.Add(lblTitle);
            this.Controls.Add(boxAdd);
            this.Controls.Add(boxList);
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (_cboCategory.SelectedItem == null) return;
            string category = _cboCategory.SelectedItem.ToString();
            string newName = _txtMenuName.Text.Trim();

            if (string.IsNullOrEmpty(newName))
            {
                MessageBox.Show("請輸入選單名稱！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!AuthManager.VerifyUser("新增自訂選單需要授權，請輸入密碼：")) return;

            string targetDb = _categoryToDbMap[category];
            DataTable dt = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
            
            // 檢查是否同名
            foreach (DataRow r in dt.Rows)
            {
                if (r["分類"].ToString() == category && r["資料表名"].ToString() == newName)
                {
                    MessageBox.Show("該分類下已經存在同名的選單！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // 建立一筆對應紀錄
            DataRow row = dt.NewRow();
            row["分類"] = category;
            row["資料庫名"] = targetDb;
            row["資料表名"] = newName;
            dt.Rows.Add(row);

            if (DataManager.BulkSaveTable("SystemConfig", "CustomMenus", dt))
            {
                // 在目標資料庫中建立基礎資料表
                string createSql = $"CREATE TABLE IF NOT EXISTS [{newName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [日期] TEXT, [備註] TEXT);";
                DataManager.InitTable(targetDb, newName, createSql);

                MessageBox.Show($"選單【{newName}】已新增至【{category}】分類下方！\n(請重新開啟系統以載入最新選單)", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _txtMenuName.Clear();
                RefreshCustomMenusList();
            }
        }

        private void BtnRename_Click(object sender, EventArgs e)
        {
            MessageBox.Show("請直接點選下方清單中的【更名】按鈕進行操作。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RefreshCustomMenusList()
        {
            _flpCustomMenus.Controls.Clear();
            DataTable dt = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");

            if (dt == null || dt.Rows.Count == 0)
            {
                _flpCustomMenus.Controls.Add(new Label { Text = "(尚無自訂選單)", ForeColor = Color.DimGray, AutoSize = true });
                return;
            }

            foreach (DataRow row in dt.Rows)
            {
                int id = Convert.ToInt32(row["Id"]);
                string category = row["分類"].ToString();
                string dbName = row["資料庫名"].ToString();
                string tableName = row["資料表名"].ToString();

                Panel pItem = new Panel { Width = 460, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(2) };
                Label lName = new Label { Text = $"[{category}] {tableName}", Location = new Point(10, 10), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                
                Button btnRename = new Button { Text = "✏️ 更名", Width = 75, Height = 35, Location = new Point(295, 5), BackColor = Color.DarkOrange, ForeColor = Color.White, Cursor = Cursors.Hand };
                btnRename.Click += (s, e) => ExecuteRename(id, dbName, tableName);

                Button btnDel = new Button { Text = "🗑️ 刪除", Width = 75, Height = 35, Location = new Point(375, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand };
                btnDel.Click += (s, e) => ExecuteDelete(id, dbName, tableName);

                pItem.Controls.Add(lName);
                pItem.Controls.Add(btnRename);
                pItem.Controls.Add(btnDel);

                _flpCustomMenus.Controls.Add(pItem);
            }
        }

        // 🟢 自製的輸入對話框，不依賴 Microsoft.VisualBasic
        private string ShowInputBox(string prompt, string title, string defaultValue)
        {
            using (Form form = new Form())
            {
                form.Width = 400;
                form.Height = 220;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.BackColor = Color.White;

                Label label = new Label() { Left = 20, Top = 20, Text = prompt, AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F) };
                TextBox textBox = new TextBox() { Left = 20, Top = 60, Width = 340, Text = defaultValue, Font = new Font("Microsoft JhengHei UI", 12F) };
                
                Button confirmation = new Button() { Text = "確認", Left = 160, Width = 90, Height = 35, Top = 120, DialogResult = DialogResult.OK, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 11F) };
                Button cancel = new Button() { Text = "取消", Left = 270, Width = 90, Height = 35, Top = 120, DialogResult = DialogResult.Cancel, Font = new Font("Microsoft JhengHei UI", 11F) };

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(confirmation);
                form.Controls.Add(cancel);
                form.AcceptButton = confirmation;

                return form.ShowDialog(this) == DialogResult.OK ? textBox.Text : "";
            }
        }

        private void ExecuteRename(int id, string dbName, string oldTableName)
        {
            if (!AuthManager.VerifyUser("更名需要授權，請輸入密碼：")) return;

            string newName = ShowInputBox($"請輸入【{oldTableName}】的新名稱：", "重新命名選單", oldTableName);
            if (string.IsNullOrWhiteSpace(newName) || newName == oldTableName) return;

            try
            {
                // 1. 更新資料庫中的真實表名
                DataManager.InitTable(dbName, newName, $"CREATE TABLE IF NOT EXISTS [{newName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT);");
                DataManager.RenameTable(dbName, oldTableName, newName);

                // 2. 更新 CustomMenus 紀錄
                DataTable dt = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
                foreach (DataRow r in dt.Rows)
                {
                    if (Convert.ToInt32(r["Id"]) == id)
                    {
                        r["資料表名"] = newName;
                        break;
                    }
                }
                DataManager.BulkSaveTable("SystemConfig", "CustomMenus", dt);
                
                MessageBox.Show("選單更名成功！(請重新開啟系統以更新畫面)", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshCustomMenusList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更名失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExecuteDelete(int id, string dbName, string tableName)
        {
            // 🟢 嚴格防護：刪除需要依序輸入 Lv2 與 Lv3
            if (!AuthManager.VerifyTableDelete()) return;

            if (MessageBox.Show($"您確定要永久刪除選單【{tableName}】及其所有資料嗎？\n(此操作無法復原)", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    DataManager.DropTable(dbName, tableName);
                    DataManager.DeleteRecord("SystemConfig", "CustomMenus", id);

                    MessageBox.Show("選單及資料已成功刪除！(請重新開啟系統以更新畫面)", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshCustomMenusList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"刪除失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
