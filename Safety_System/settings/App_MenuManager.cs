/// FILE: Safety_System/settings/App_MenuManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_MenuManager : Form
    {
        private ComboBox _cboCategory;
        private TextBox _txtMenuName;
        private Button _btnAdd;
        
        private FlowLayoutPanel _flpCustomMenus;

        private readonly Dictionary<string, string> _categoryToDbMap = new Dictionary<string, string>
        {
            { "日常作業", "Reports" },
            { "工安", "Safety" },
            { "化學品", "Chemical" },
            { "化學品要求及規範", "Chemical" },
            { "護理", "Nursing" },
            { "空污", "Air" },
            { "水污", "Water" },
            { "廢棄物", "Waste" },
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
            string createSql = "CREATE TABLE IF NOT EXISTS [CustomMenus] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [分類] TEXT, [資料庫名] TEXT, [資料表名] TEXT);";
            DataManager.InitTable("SystemConfig", "CustomMenus", createSql);
            
            InitializeComponent();
            RefreshCustomMenusList();
        }

        private void InitializeComponent()
        {
            this.Text = "選單新增與管理";
            this.Size = new Size(650, 750);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label { Text = "📂 自訂選單新增與管理", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.SteelBlue, Location = new Point(25, 20), AutoSize = true };

            GroupBox boxAdd = new GroupBox { Text = "操作區 (新增)", Location = new Point(25, 70), Size = new Size(590, 160), Font = new Font("Microsoft JhengHei UI", 12F) };

            Label lblCat = new Label { Text = "目標分類：", Location = new Point(20, 45), AutoSize = true };
            _cboCategory = new ComboBox { Location = new Point(150, 42), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            _cboCategory.Items.AddRange(_categoryToDbMap.Keys.ToArray());
            if (_cboCategory.Items.Count > 0) _cboCategory.SelectedIndex = 0;

            Label lblName = new Label { Text = "選單名稱：", Location = new Point(20, 100), AutoSize = true };
            _txtMenuName = new TextBox { Location = new Point(150, 97), Width = 220 };

            _btnAdd = new Button { Text = "➕ 新增", Location = new Point(410, 38), Size = new Size(140, 90), BackColor = Color.ForestGreen, ForeColor = Color.White, Cursor = Cursors.Hand };
            _btnAdd.Click += BtnAdd_Click;

            boxAdd.Controls.Add(lblCat);
            boxAdd.Controls.Add(_cboCategory);
            boxAdd.Controls.Add(lblName);
            boxAdd.Controls.Add(_txtMenuName);
            boxAdd.Controls.Add(_btnAdd);

            GroupBox boxList = new GroupBox { Text = "已建立之自訂選單 (更名 / 刪除)", Location = new Point(25, 250), Size = new Size(590, 430), Font = new Font("Microsoft JhengHei UI", 12F) };
            
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

            bool isPersonalMenu = category.StartsWith("選單");
            if (isPersonalMenu)
            {
                if (!AuthManager.VerifyHiddenMenu(category)) return;
            }

            string authPrompt = "新增自訂選單需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            string targetDb = _categoryToDbMap[category];
            DataTable dt = DataManager.GetTableData("SystemConfig", "CustomMenus", "", "", "");
            
            foreach (DataRow r in dt.Rows)
            {
                if (r["分類"].ToString() == category && r["資料表名"].ToString() == newName)
                {
                    MessageBox.Show("該分類下已經存在同名的選單！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var existingCols = DataManager.GetColumnNames(targetDb, newName);
            if (existingCols.Count > 0)
            {
                MessageBox.Show($"資料庫中已經存在名為【{newName}】的資料表！請更換其他名稱。", "命名衝突", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataRow row = dt.NewRow();
            row["分類"] = category;
            row["資料庫名"] = targetDb;
            row["資料表名"] = newName;
            dt.Rows.Add(row);

            if (DataManager.BulkSaveTable("SystemConfig", "CustomMenus", dt))
            {
                string schema = TableSchemaManager.DefaultCustomSchema;
                string createSql = $"CREATE TABLE IF NOT EXISTS [{newName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, {schema});";
                DataManager.InitTable(targetDb, newName, createSql);

                MessageBox.Show($"選單【{newName}】已新增至【{category}】分類下方！\n(請重新開啟系統以載入最新選單)", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _txtMenuName.Clear();
                RefreshCustomMenusList();
            }
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

            var groupedMenus = dt.Rows.Cast<DataRow>()
                                 .GroupBy(r => r["分類"].ToString())
                                 .OrderBy(g => g.Key);

            foreach (var group in groupedMenus)
            {
                string category = group.Key;
                bool isPersonalMenu = category.StartsWith("選單");

                Panel pnlHeader = new Panel { Width = 540, Height = 40, BackColor = Color.LightSteelBlue, Margin = new Padding(0, 5, 0, 0) };
                
                Button btnToggle = new Button { 
                    Text = isPersonalMenu ? "+" : "-", 
                    Width = 35, 
                    Height = 35, 
                    Location = new Point(5, 2), 
                    Font = new Font("Consolas", 14F, FontStyle.Bold), 
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.WhiteSmoke,
                    Cursor = Cursors.Hand
                };
                btnToggle.FlatAppearance.BorderSize = 0;

                Label lblCat = new Label { 
                    Text = $"📂 {category} ({group.Count()} 個選單)", 
                    Location = new Point(50, 8), 
                    AutoSize = true, 
                    Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                    ForeColor = Color.DarkSlateBlue
                };

                pnlHeader.Controls.Add(btnToggle);
                pnlHeader.Controls.Add(lblCat);

                FlowLayoutPanel flpItems = new FlowLayoutPanel { 
                    Width = 540, 
                    AutoSize = true, 
                    FlowDirection = FlowDirection.TopDown, 
                    WrapContents = false, 
                    Margin = new Padding(0, 0, 0, 10),
                    Visible = !isPersonalMenu 
                };

                btnToggle.Click += (s, e) => {
                    if (flpItems.Visible) {
                        flpItems.Visible = false;
                        btnToggle.Text = "+";
                    } else {
                        if (isPersonalMenu) {
                            if (!AuthManager.VerifyHiddenMenu(category)) return;
                        }
                        flpItems.Visible = true;
                        btnToggle.Text = "-";
                    }
                };

                foreach (DataRow row in group)
                {
                    int id = Convert.ToInt32(row["Id"]);
                    string dbName = row["資料庫名"].ToString();
                    string tableName = row["資料表名"].ToString();

                    Panel pItem = new Panel { Width = 520, Height = 45, BackColor = Color.WhiteSmoke, Margin = new Padding(20, 2, 0, 2) };
                    Label lName = new Label { Text = $"• {tableName}", Location = new Point(10, 12), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) };
                    
                    Button btnRename = new Button { Text = "✏️ 更名", Width = 80, Height = 35, Location = new Point(340, 5), BackColor = Color.DarkOrange, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                    btnRename.FlatAppearance.BorderSize = 0;
                    btnRename.Click += (s, e) => ExecuteRename(id, dbName, tableName);

                    Button btnDel = new Button { Text = "🗑️ 刪除", Width = 80, Height = 35, Location = new Point(430, 5), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
                    btnDel.FlatAppearance.BorderSize = 0;
                    btnDel.Click += (s, e) => ExecuteDelete(id, dbName, tableName);

                    pItem.Controls.Add(lName);
                    pItem.Controls.Add(btnRename);
                    pItem.Controls.Add(btnDel);

                    flpItems.Controls.Add(pItem);
                }

                _flpCustomMenus.Controls.Add(pnlHeader);
                _flpCustomMenus.Controls.Add(flpItems);
            }
        }

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
            string authPrompt = "更名自訂選單需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：";
            if (!AuthManager.VerifyAdmin(authPrompt)) return;

            string newName = ShowInputBox($"請輸入【{oldTableName}】的新名稱：", "重新命名選單", oldTableName);
            if (string.IsNullOrWhiteSpace(newName) || newName == oldTableName) return;

            try
            {
                var existingCols = DataManager.GetColumnNames(dbName, newName);
                if (existingCols.Count > 0)
                {
                    MessageBox.Show($"資料庫中已經存在名為【{newName}】的資料表！請更換其他名稱。", "命名衝突", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataManager.RenameTable(dbName, oldTableName, newName);

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
