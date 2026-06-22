/// FILE: Safety_System/settings/App_DbConfig.Management.cs ///
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
        private void BuildManagementTab(TabPage tabDb)
        {
            Panel pnlDb = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

            // =====================================
            // 1. 資料庫 (DB) 存放路徑設定
            // =====================================
            GroupBox boxPath = new GroupBox { Text = "資料庫 (DB) 存放路徑設定", Dock = DockStyle.Top, Height = 180, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            
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
                if (System.IO.Directory.Exists(_txtPath.Text)) {
                    DataManager.SetBasePath(_txtPath.Text);
                    MessageBox.Show("DB 路徑已更新！後續系統存取皆會依此路徑。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("請選擇有效的資料夾路徑。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            boxPath.Controls.AddRange(new Control[] { _txtPath, btnBrowse, btnSavePath });

            // =====================================
            // 2. 附件檔案存放路徑設定
            // =====================================
            GroupBox boxAttachPath = new GroupBox { Text = "附件檔案存放路徑設定", Dock = DockStyle.Top, Height = 180, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Padding = new Padding(15) };
            boxAttachPath.Margin = new Padding(0, 30, 0, 0);

            string currentAttachPath = string.IsNullOrEmpty(DataManager.AttachmentBasePath) ? "" : DataManager.AttachmentBasePath;
            _txtAttachmentPath = new TextBox { Location = new Point(30, 50), Width = 600, ReadOnly = true, Text = currentAttachPath, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Button btnBrowseAttach = new Button { Text = "選擇資料夾", Location = new Point(650, 48), Size = new Size(150, 35), Font = new Font("Microsoft JhengHei UI", 12F) };
            btnBrowseAttach.Click += (s, e) => {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "請選擇附件檔案存放的資料夾" }) {
                    if (fbd.ShowDialog() == DialogResult.OK) _txtAttachmentPath.Text = fbd.SelectedPath;
                }
            };
            
            Button btnSaveAttachPath = new Button { Text = "儲存附件路徑變更", Location = new Point(30, 110), Size = new Size(220, 45), BackColor = Color.Teal, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnSaveAttachPath.Click += (s, e) => {
                if (System.IO.Directory.Exists(_txtAttachmentPath.Text)) {
                    DataManager.SetAttachmentBasePath(_txtAttachmentPath.Text);
                    MessageBox.Show("附件路徑已更新！後續系統存取皆會依此路徑。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    MessageBox.Show("請選擇有效的資料夾路徑。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            boxAttachPath.Controls.AddRange(new Control[] { _txtAttachmentPath, btnBrowseAttach, btnSaveAttachPath });

            // =====================================
            // 3. 資料庫備份設定
            // =====================================
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
                BackupManager.SaveConfig(_txtBackupPath.Text, (int)_numKeepCount.Value, (int)_numIntervalDays.Value);
                MessageBox.Show("備份設定已儲存！", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Button btnManualBackup = new Button { Text = "立即執行手動熱備份", Location = new Point(300, 220), Size = new Size(220, 45), BackColor = Color.DimGray, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F) };
            btnManualBackup.Click += (s, e) => {
                BackupManager.ExecuteBackup();
                MessageBox.Show("熱備份(Hot Backup)執行完成！\n不會影響目前操作中的使用者。", "備份成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            boxBackup.Controls.AddRange(new Control[] { lblB1, _txtBackupPath, btnBrowseBackup, lblB2, _numKeepCount, lblB3, lblB4, _numIntervalDays, lblB5, btnSaveBackup, btnManualBackup });

            // =====================================
            // 4. 強制刪除資料表 (危險操作)
            // =====================================
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

            pnlDb.Controls.Add(boxDelete);
            pnlDb.Controls.Add(spacer3);
            pnlDb.Controls.Add(boxBackup);
            pnlDb.Controls.Add(spacer2);
            pnlDb.Controls.Add(boxAttachPath); 
            pnlDb.Controls.Add(spacer1);
            pnlDb.Controls.Add(boxPath);

            tabDb.Controls.Add(pnlDb);
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

        private void BuildAuditTab(TabPage tabAudit)
        {
            Panel pnlAudit = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(20) };

            GroupBox boxAudit = new GroupBox { Text = "操作軌跡追蹤 (查閱最後修改人與時間)", Dock = DockStyle.Top, Height = 650, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.DarkSlateBlue, Padding = new Padding(15) };

            Label lblAuditDb = new Label { Text = "選擇資料庫:", Location = new Point(30, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.Black };
            _cboAuditDb = new ComboBox { Location = new Point(150, 58), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            Label lblAuditTable = new Label { Text = "選擇資料表:", Location = new Point(360, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F), ForeColor = Color.Black };
            _cboAuditTable = new ComboBox { Location = new Point(480, 58), Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };

            _chkShowDeletedLogs = new CheckBox { Text = "☑️ 僅查詢該表「被刪除的資料」軌跡", Location = new Point(760, 60), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 10F), ForeColor = Color.Crimson, Cursor = Cursors.Hand };

            Button btnSearchAudit = new Button { Text = "🔍 查詢操作紀錄", Location = new Point(760, 110), Size = new Size(180, 35), BackColor = Color.DarkSlateBlue, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
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
            
            typeof(DataGridView).InvokeMember("DoubleBuffered", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, 
                null, _dgvAudit, new object[] { true });

            boxAudit.Controls.AddRange(new Control[] { lblAuditDb, _cboAuditDb, lblAuditTable, _cboAuditTable, _chkShowDeletedLogs, btnSearchAudit, _dgvAudit });

            pnlAudit.Controls.Add(boxAudit);
            tabAudit.Controls.Add(pnlAudit);
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
                    // 🟢 核心修復：統一吃 SysConfigDbPath
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
    }
}
