/// FILE: Safety_System/App_ChemDashboard.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_ChemDashboard
    {
        // UI 控制項成員
        private DataGridView _dgvSDS;
        
        // 常數定義
        private const string DbName = "Chemical";
        private const string TableName = "SDS_Inventory";
        
        // 記憶設定檔路徑
        private readonly string VisibilityFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChemSDS_Visibility.txt");
        
        // 用於存儲欄位顯示狀態的字典 (Key: 欄位名稱, Value: 是否顯示)
        private Dictionary<string, bool> _columnVisibility = new Dictionary<string, bool>();

        /// <summary>
        /// 進入模組的主入口，回傳主畫面控制項
        /// </summary>
        public Control GetView()
        {
            // 1. 啟動時先載入使用者先前的顯示偏好設定
            LoadVisibilitySettings();

            // 2. 建立主容器排版 (兩行：第一行為按鈕，第二行為內容)
            TableLayoutPanel mainLayout = new TableLayoutPanel { 
                Dock = DockStyle.Fill, 
                Padding = new Padding(20), 
                RowCount = 2,
                BackColor = Color.WhiteSmoke
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 3. 第一行：功能按鈕區域 (FlowLayoutPanel)
            FlowLayoutPanel pnlAction = new FlowLayoutPanel { 
                Dock = DockStyle.Fill, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, 15) 
            };
            
            // 按鈕 A：PDF 導出
            Button btnPdf = new Button { 
                Text = "📤 導出全廠 SDS 清冊 PDF", 
                Size = new Size(220, 45), 
                BackColor = Color.DarkCyan, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand 
            };
            btnPdf.Click += (s, e) => ExportToPdf();

            // 按鈕 B：欄位顯示設定
            Button btnSettings = new Button { 
                Text = "⚙️ 設定顯示欄位", 
                Size = new Size(200, 45), 
                BackColor = Color.LightSlateGray, 
                ForeColor = Color.White, 
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                Cursor = Cursors.Hand, 
                Margin = new Padding(15, 0, 0, 0) 
            };
            btnSettings.Click += (s, e) => OpenColumnSettings();

            pnlAction.Controls.Add(btnPdf);
            pnlAction.Controls.Add(btnSettings);
            mainLayout.Controls.Add(pnlAction, 0, 0);

            // 4. 第二行：主看板大框 (GroupBox)
            GroupBox boxMain = new GroupBox { 
                Text = "📋 化學品管理綜合看板", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                Padding = new Padding(15) 
            };
            
            TableLayoutPanel innerTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1 };
            innerTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F)); // 小框 1：標題高度
            innerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 小框 2：表格高度

            // --- 小框 1：企業標題與裝飾 ---
            Panel pnlTitle = new Panel { 
                Dock = DockStyle.Fill, 
                BackColor = Color.FromArgb(240, 245, 250), 
                Margin = new Padding(0, 0, 0, 15) 
            };
            pnlTitle.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, pnlTitle.ClientRectangle, Color.SteelBlue, ButtonBorderStyle.Solid);
            
            Label lblCompany = new Label { 
                Text = "台灣玻璃工業股份有限公司 - 彰濱廠", 
                Dock = DockStyle.Top, 
                Height = 45, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold), 
                ForeColor = Color.MidnightBlue 
            };
            Label lblSubTitle = new Label { 
                Text = "化學品清單一覽表 (SDS 定期追蹤管理)", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleCenter, 
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold), 
                ForeColor = Color.FromArgb(60, 60, 60) 
            };
            pnlTitle.Controls.Add(lblSubTitle);
            pnlTitle.Controls.Add(lblCompany);

            // --- 小框 2：SDS 數據表格區域 ---
            GroupBox boxGrid = new GroupBox { 
                Text = "SDS 安全資料表庫存明細", 
                Dock = DockStyle.Fill, 
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold) 
            };
            _dgvSDS = new DataGridView { 
                Dock = DockStyle.Fill, 
                BackgroundColor = Color.White, 
                AllowUserToAddRows = false, 
                ReadOnly = true, 
                RowHeadersVisible = false, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                BorderStyle = BorderStyle.None 
            };
            _dgvSDS.RowTemplate.Height = 35;
            _dgvSDS.EnableHeadersVisualStyles = false;
            _dgvSDS.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 62, 80);
            _dgvSDS.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvSDS.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold);
            _dgvSDS.ColumnHeadersHeight = 40;
            _dgvSDS.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
            boxGrid.Controls.Add(_dgvSDS);

            // 組裝畫面
            innerTable.Controls.Add(pnlTitle, 0, 0);
            innerTable.Controls.Add(boxGrid, 0, 1);
            boxMain.Controls.Add(innerTable);
            mainLayout.Controls.Add(boxMain, 0, 1);

            // 5. 嘗試讀取資料 (即使資料庫尚未建立也不會當機)
            LoadData();

            return mainLayout;
        }

        /// <summary>
        /// 從資料庫抓取 SDS 資料
        /// </summary>
        private void LoadData()
        {
            try
            {
                DataTable dt = DataManager.GetTableData(DbName, TableName, "", "", "");
                if (dt != null && dt.Columns.Count > 0)
                {
                    _dgvSDS.DataSource = dt;
                    ApplyVisibility(); // 立即套用欄位隱藏/顯示
                }
                else
                {
                    _dgvSDS.DataSource = null;
                }
            }
            catch { _dgvSDS.DataSource = null; }
        }

        /// <summary>
        /// 根據使用者的偏好字典 ApplyVisibility 到 DataGridView
        /// </summary>
        private void ApplyVisibility()
        {
            if (_dgvSDS.Columns.Count == 0) return;
            foreach (DataGridViewColumn col in _dgvSDS.Columns)
            {
                // 固定隱藏 Id 欄位
                if (col.Name.Equals("Id", StringComparison.OrdinalIgnoreCase)) { col.Visible = false; continue; }
                
                if (_columnVisibility.ContainsKey(col.Name))
                {
                    col.Visible = _columnVisibility[col.Name];
                }
                else
                {
                    // 若無記憶設定，則預設顯示 (排除附件檔案字串)
                    col.Visible = !col.Name.Contains("附件");
                }
            }
        }

        /// <summary>
        /// 彈出對話框供使用者動態勾選欲顯示的欄位
        /// </summary>
        private void OpenColumnSettings()
        {
            try
            {
                // 從資料庫元數據獲取所有最新的欄位名稱
                List<string> allCols = DataManager.GetColumnNames(DbName, TableName);
                if (allCols == null || allCols.Count == 0)
                {
                    MessageBox.Show("目前找不到資料表，請先至「資料庫設定」初始化模組或匯入資料。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (Form f = new Form())
                {
                    f.Text = "⚙️ 看板顯示欄位設定";
                    f.Size = new Size(380, 550);
                    f.StartPosition = FormStartPosition.CenterParent;
                    f.FormBorderStyle = FormBorderStyle.FixedDialog;
                    f.MaximizeBox = false; f.MinimizeBox = false;
                    f.BackColor = Color.White;

                    Label lbl = new Label { 
                        Text = "請勾選要在看板與報表中顯示的項目：", 
                        Dock = DockStyle.Top, 
                        Height = 50, 
                        Padding = new Padding(10), 
                        Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                        ForeColor = Color.SteelBlue
                    };
                    
                    CheckedListBox clb = new CheckedListBox { 
                        Dock = DockStyle.Fill, 
                        CheckOnClick = true, 
                        Font = new Font("Microsoft JhengHei UI", 11F),
                        BorderStyle = BorderStyle.None,
                        BackColor = Color.FromArgb(250, 250, 250)
                    };
                    
                    // 填入清單
                    foreach (var colName in allCols)
                    {
                        if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase)) continue;
                        
                        // 決定預設勾選狀態
                        bool isChecked = true;
                        if (_columnVisibility.ContainsKey(colName)) isChecked = _columnVisibility[colName];
                        else if (colName.Contains("附件")) isChecked = false;

                        clb.Items.Add(colName, isChecked);
                    }

                    // 儲存按鈕
                    Button btnSave = new Button { 
                        Text = "💾 儲存設定並套用", 
                        Dock = DockStyle.Bottom, 
                        Height = 55, 
                        BackColor = Color.ForestGreen, 
                        ForeColor = Color.White, 
                        Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), 
                        Cursor = Cursors.Hand,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnSave.Click += (s, e) => {
                        _columnVisibility.Clear();
                        for (int i = 0; i < clb.Items.Count; i++)
                        {
                            _columnVisibility[clb.Items[i].ToString()] = clb.GetItemChecked(i);
                        }
                        SaveVisibilitySettings();
                        ApplyVisibility();
                        f.DialogResult = DialogResult.OK;
                    };

                    f.Controls.Add(clb);
                    f.Controls.Add(lbl);
                    f.Controls.Add(btnSave);
                    f.ShowDialog();
                }
            }
            catch (Exception ex) { MessageBox.Show("開啟設定視窗失敗：" + ex.Message); }
        }

        /// <summary>
        /// 從 TXT 讀取欄位顯示紀錄
        /// </summary>
        private void LoadVisibilitySettings()
        {
            _columnVisibility.Clear();
            if (File.Exists(VisibilityFile))
            {
                try
                {
                    string[] lines = File.ReadAllLines(VisibilityFile, Encoding.UTF8);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('|');
                        if (parts.Length == 2)
                        {
                            _columnVisibility[parts[0]] = (parts[1] == "1");
                        }
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 將欄位顯示紀錄寫入 TXT
        /// </summary>
        private void SaveVisibilitySettings()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (var kvp in _columnVisibility)
                {
                    sb.AppendLine($"{kvp.Key}|{(kvp.Value ? "1" : "0")}");
                }
                File.WriteAllText(VisibilityFile, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        /// <summary>
        /// 完整分頁導出 PDF 報表引擎
        /// </summary>
        private void ExportToPdf()
        {
            if (_dgvSDS.DataSource == null || _dgvSDS.Rows.Count == 0)
            {
                MessageBox.Show("目前看板中沒有數據可供導出。");
                return;
            }

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true; // A4 橫向列印以容納更多欄位
            pd.DefaultPageSettings.Margins = new Margins(30, 30, 30, 30);
            
            int rowIndex = 0; // 用於追蹤目前的列印進度 (分頁用)

            pd.PrintPage += (s, e) => {
                Graphics g = e.Graphics;
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                // A. 繪製報表標頭
                Font fTitle = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold);
                Font fSub = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold);
                Font fBody = new Font("Microsoft JhengHei UI", 9F);
                Font fHead = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);

                g.DrawString("台灣玻璃工業股份有限公司 - 彰濱廠", fTitle, Brushes.MidnightBlue, x, y);
                y += 40;
                g.DrawString("化學品清單一覽表 (SDS)", fSub, Brushes.Black,
