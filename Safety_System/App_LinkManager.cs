/// FILE: Safety_System/App_LinkManager.cs ///
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_LinkManager : Form
    {
        private TextBox _txtName;
        private TextBox _txtPath;
        private DataGridView _dgv;
        private const string DbName = "SystemConfig";
        private const string TableName = "AppLinks";

        public App_LinkManager()
        {
            // 在畫面初始化前，強制確保資料庫與資料表結構存在，防止開啟或新增時崩潰
            string createSql = "CREATE TABLE IF NOT EXISTS [AppLinks] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [選單名稱] TEXT, [執行路徑] TEXT);";
            DataManager.InitTable(DbName, TableName, createSql);

            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "應用連結設定 (可自訂外部程式快捷鍵)";
            this.Size = new Size(650, 560);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            Label lblTitle = new Label { Text = "🔗 應用選單連結設定", Font = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold), ForeColor = Color.SteelBlue, Location = new Point(20, 20), AutoSize = true };

            GroupBox boxAdd = new GroupBox { Text = "新增外部程式連結", Location = new Point(25, 70), Size = new Size(580, 140) };
            
            Label lblName = new Label { Text = "選單名稱：", Location = new Point(20, 40), AutoSize = true };
            _txtName = new TextBox { Location = new Point(160, 37), Width = 270 };

            Label lblPath = new Label { Text = "執行路徑 (.exe)：", Location = new Point(20, 85), AutoSize = true };
            _txtPath = new TextBox { Location = new Point(160, 82), Width = 270 };

            Button btnBrowse = new Button { Text = "瀏覽", Location = new Point(440, 80), Size = new Size(115, 33), Font = new Font("Microsoft JhengHei UI", 11F), Cursor = Cursors.Hand };
            btnBrowse.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "執行檔 (*.exe)|*.exe|所有檔案 (*.*)|*.*", Title = "選擇外部應用程式" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) _txtPath.Text = ofd.FileName;
                }
            };

            Button btnAdd = new Button { Text = "➕ 新增", Location = new Point(440, 35), Size = new Size(115, 38), BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnAdd.FlatAppearance.BorderSize = 0;
            btnAdd.Click += BtnAdd_Click;

            boxAdd.Controls.AddRange(new Control[] { lblName, _txtName, lblPath, _txtPath, btnBrowse, btnAdd });

            _dgv = new DataGridView {
                Location = new Point(25, 230), Size = new Size(580, 215), 
                BackgroundColor = Color.WhiteSmoke, AllowUserToAddRows = false, ReadOnly = true,
                // 🟢 修正：將 Fill 改為 DisplayedCells，文字長度超過視窗時就會自動產生水平捲軸
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, 
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            Button btnDel = new Button { Text = "🗑️ 刪除選取連結", Location = new Point(25, 460), Size = new Size(180, 40), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand, FlatStyle = FlatStyle.Flat };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += BtnDel_Click;

            this.Controls.AddRange(new Control[] { lblTitle, boxAdd, _dgv, btnDel });
        }

        private void LoadData()
        {
            try
            {
                DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
                
                if (dt != null && dt.Columns.Count == 0)
                {
                    dt.Columns.Add("Id", typeof(int));
                    dt.Columns.Add("選單名稱", typeof(string));
                    dt.Columns.Add("執行路徑", typeof(string));
                }

                _dgv.DataSource = dt;
                if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].Visible = false;
                
                // 🟢 微調：確保即使名稱很短，也不會太擠，同時路徑保持自動延展
                if (_dgv.Columns.Contains("選單名稱")) _dgv.Columns["選單名稱"].MinimumWidth = 150;

                _dgv.ClearSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取資料失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string name = _txtName.Text.Trim();
            string path = _txtPath.Text.Trim();

            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(path)) {
                MessageBox.Show("名稱與路徑不可空白！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
                
                if (dt.Columns.Count == 0)
                {
                    dt.Columns.Add("Id", typeof(int));
                    dt.Columns.Add("選單名稱", typeof(string));
                    dt.Columns.Add("執行路徑", typeof(string));
                }

                DataRow r = dt.NewRow();
                r["選單名稱"] = name;
                r["執行路徑"] = path;
                dt.Rows.Add(r);

                if (DataManager.BulkSaveTable(DbName, TableName, dt)) {
                    _txtName.Clear();
                    _txtPath.Clear();
                    LoadData();
                    MessageBox.Show("新增成功！關閉視窗後選單即會更新。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增儲存失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (_dgv.SelectedRows.Count == 0) return;

            if (MessageBox.Show("確定要刪除選取的連結嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                try
                {
                    int id = Convert.ToInt32(_dgv.SelectedRows[0].Cells["Id"].Value);
                    DataManager.DeleteRecord(DbName, TableName, id);
                    LoadData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("刪除失敗：" + ex.Message, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
