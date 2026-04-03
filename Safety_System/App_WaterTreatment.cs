using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_WaterTreatment
    {
        private DataGridView _dgv;
        private DateTimePicker _dtpStart, _dtpEnd;
        private TextBox _txtNewColName;

        public Control GetView()
        {
            // 主容器：使用 TableLayoutPanel 分成上下兩個大框
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 180F)); // 上框高度
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 下框滿版

            // --- 第一個框：操作與設定 ---
            GroupBox boxTop = new GroupBox { Text = "控制與新增欄位", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            // 第一行：讀取功能
            Label lblDate = new Label { Text = "日期區間:", Location = new Point(20, 40), AutoSize = true };
            _dtpStart = new DateTimePicker { Location = new Point(110, 35), Width = 150, Format = DateTimePickerFormat.Short };
            Label lblTo = new Label { Text = "至", Location = new Point(270, 40), AutoSize = true };
            _dtpEnd = new DateTimePicker { Location = new Point(300, 35), Width = 150, Format = DateTimePickerFormat.Short };
            
            Button btnRead = new Button { Text = "讀取資料", Location = new Point(470, 32), Size = new Size(120, 35), BackColor = Color.LightBlue };
            btnRead.Click += BtnRead_Click;

            // 第二行：新增欄位功能
            Label lblNewCol = new Label { Text = "新增欄位標題:", Location = new Point(20, 100), AutoSize = true };
            _txtNewColName = new TextBox { Location = new Point(150, 95), Width = 200 };
            Button btnAddCol = new Button { Text = "確認新增欄位", Location = new Point(370, 92), Size = new Size(150, 35), BackColor = Color.LightGray };
            btnAddCol.Click += BtnAddCol_Click;

            boxTop.Controls.AddRange(new Control[] { lblDate, _dtpStart, lblTo, _dtpEnd, btnRead, lblNewCol, _txtNewColName, btnAddCol });
            mainLayout.Controls.Add(boxTop, 0, 0);

            // --- 第二個框：數據顯示與即時編輯 ---
            GroupBox boxBottom = new GroupBox { Text = "水處理數據明細 (即時存檔模式)", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
            _dgv = new DataGridView {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                AllowUserToAddRows = true // 允許在最下方直接新增空行
            };
            
            // 綁定事件：當儲存格修改完成後立即存檔
            _dgv.CellValueChanged += Dgv_CellValueChanged;
            
            boxBottom.Controls.Add(_dgv);
            mainLayout.Controls.Add(boxBottom, 0, 1);

            return mainLayout;
        }

        private void BtnRead_Click(object sender, EventArgs e)
        {
            // 檢查資料庫是否存在
            if (!DataManager.IsDbFileExists())
            {
                var result = MessageBox.Show("未找到資料庫檔案，是否立即建立新的水處理資料表？", "系統提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DataManager.CreateWaterTable();
                    MessageBox.Show("資料表已建立成功！");
                }
                else return;
            }

            RefreshGrid();
        }

        private void RefreshGrid()
        {
            string s = _dtpStart.Value.ToString("yyyy-MM-dd");
            string e = _dtpEnd.Value.ToString("yyyy-MM-dd");
            _dgv.DataSource = DataManager.GetWaterData(s, e);
            
            // 隱藏 ID 欄位不讓使用者改
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
        }

        private void BtnAddCol_Click(object sender, EventArgs e)
        {
            string colName = _txtNewColName.Text.Trim();
            if (string.IsNullOrEmpty(colName)) return;

            try {
                DataManager.AddColumnToWaterTable(colName);
                MessageBox.Show("欄位 [" + colName + "] 新增成功！請重新點擊讀取以載入新架構。");
                _txtNewColName.Clear();
            }
            catch (Exception ex) {
                MessageBox.Show("欄位新增失敗：" + ex.Message);
            }
        }

        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // 排除標題行與初始化階段
            if (e.RowIndex < 0) return;

            var row = _dgv.Rows[e.RowIndex];
            var idVal = row.Cells["Id"].Value;

            if (idVal != null && idVal != DBNull.Value)
            {
                int id = Convert.ToInt32(idVal);
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                object newVal = row.Cells[e.ColumnIndex].Value;

                // 執行即時更新
                DataManager.UpdateWaterCell(id, colName, newVal);
            }
        }
    }
}
