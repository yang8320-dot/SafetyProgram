/// FILE: Safety_System/settings/App_DbConfig.Keys.cs ///
using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public partial class App_DbConfig
    {
        private void BuildKeysTab(TabPage tabKeys)
        {
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

            boxKeys.Controls.AddRange(new Control[] { lblDb, _cboDb, lblTable, _cboTable, lblCol1, _cboCol1, lblCol2, _cboCol2, lblCol3, _cboCol3, lblCol4, _cboCol4, btnSaveKeys, btnShowTableKeysList });

            pnlKeys.Controls.Add(boxKeys);
            tabKeys.Controls.Add(pnlKeys);
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

        private void BtnShowTableKeysList_Click(object sender, EventArgs e)
        {
            using (Form f = new Form { Text = "防重寫設定清單", Size = new Size(900, 550), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.Sizable, MaximizeBox = true, MinimizeBox = false, BackColor = Color.White })
            {
                TableLayoutPanel tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize)); tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                Label lbl = new Label { Text = "已設定之資料表防重寫規則 (每個資料表僅限一組)：", Dock = DockStyle.Fill, Padding = new Padding(15, 15, 15, 10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true };
                tlp.Controls.Add(lbl, 0, 0);

                FlowLayoutPanel flp = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10) };
                tlp.Controls.Add(flp, 0, 1); f.Controls.Add(tlp);

                Action loadKeys = null;
                loadKeys = () => {
                    flp.Controls.Clear();
                    try {
                        string sysDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemConfig.sqlite");
                        DataTable dt = new DataTable();
                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                            conn.Open();
                            using (var cmd = new SQLiteCommand("SELECT * FROM TableKeys", conn))
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                        }

                        if (dt.Rows.Count == 0) {
                            flp.Controls.Add(new Label { Text = "目前沒有任何防重寫設定。", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Microsoft JhengHei UI", 11F) }); return;
                        }

                        foreach (DataRow row in dt.Rows) {
                            string db = row["DbName"].ToString(); string tb = row["TableName"].ToString();
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
                                    if (AuthManager.VerifyAdmin("刪除防重寫設定需要系統權限\n請輸入【Lv2管理者】等級以上\n密碼進行授權：")) {
                                        using (var conn = new SQLiteConnection($"Data Source={sysDbPath};Version=3;")) {
                                            conn.Open();
                                            using (var cmd = new SQLiteCommand("DELETE FROM TableKeys WHERE DbName=@DB AND TableName=@TB", conn)) {
                                                cmd.Parameters.AddWithValue("@DB", db); cmd.Parameters.AddWithValue("@TB", tb); cmd.ExecuteNonQuery();
                                            }
                                        }
                                        loadKeys();
                                    }
                                }
                            };
                            p.Controls.Add(lTxt); p.Controls.Add(btnDel); flp.Controls.Add(p);
                        }
                    } catch { }
                };

                loadKeys();
                f.ShowDialog();
            }
        }
    }
}
