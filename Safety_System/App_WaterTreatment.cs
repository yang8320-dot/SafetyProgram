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
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 200F)); 
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // --- 第一個框：控制與設定 ---
            GroupBox boxTop = new GroupBox { Text = "數據檢索與欄位管理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F) };
            
            // 1. 日期區間與讀取
            Label lblDate = new Label { Text = "日期區間:", Location = new Point(20, 40), AutoSize = true };
            _dtpStart = new DateTimePicker { Location = new Point(110, 35), Width = 150, Format = DateTimePickerFormat.Short };
            
            // 預設為 30 天前
            _dtpStart.Value = DateTime.Now.AddDays(-30); 
            
            Label lblTo = new Label { Text = "至", Location = new Point(270, 40), AutoSize = true };
            _dtpEnd = new DateTimePicker { Location = new Point(300, 35), Width = 150, Format = DateTimePickerFormat.Short };
            
            Button btnRead = new Button { Text = "讀取資料庫", Location = new Point(470, 32), Size = new Size(120, 35), BackColor = Color.LightBlue };
            btnRead.Click += BtnRead_Click;

            // 2. 新增欄位功能
            Label lblNewCol = new Label { Text = "新增欄位標題:", Location = new Point(20, 90), AutoSize = true };
            _txtNewColName = new TextBox { Location = new Point(150, 85), Width = 200 };
            Button btnAddCol = new Button { Text = "確認新增欄位", Location = new Point(370, 82), Size = new Size(150, 35), BackColor = Color.LightGray };
            btnAddCol.Click += BtnAddCol_Click;

            // 3. 手動存入按鍵 (縮短寬度以騰出空間)
            Button btnSaveManual = new Button { 
                Text = "💾 手動儲存所有變更", 
                Location = new Point(20, 140), 
                Size = new Size(340, 40), 
                BackColor = Color.ForestGreen, 
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnSaveManual.Click += BtnSaveManual_Click;

            // 🟢 4. 新增：刪除選取資料按鍵
            Button btnDeleteRow = new Button {
                Text = "🗑️ 刪除選取資料",
                Location = new Point(370, 140),
                Size = new Size(150, 40),
                BackColor = Color.IndianRed,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold)
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;

            boxTop.Controls.AddRange(new Control[] { lblDate, _dtpStart, lblTo, _dtpEnd, btnRead, lblNewCol, _txtNewColName, btnAddCol, btnSaveManual, btnDeleteRow });
            mainLayout.Controls.Add(boxTop, 0, 0);

            // --- 第二個框：數據顯示 ---
            GroupBox boxBottom = new GroupBox { Text = "水處理數據明細", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 11F) };
            _dgv = new DataGridView {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                SelectionMode = DataGridViewSelectionMode.RowHeaderSelect, // 方便整列選取
                AllowUserToAddRows = true 
            };
            
            boxBottom.Controls.Add(_dgv);
            mainLayout.Controls.Add(boxBottom, 0, 1);

            return mainLayout;
        }

        private void BtnRead_Click(object sender, EventArgs e)
        {
            if (!DataManager.IsDbFileExists())
            {
                if (MessageBox.Show("未找到資料庫，是否新增？", "詢問", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    DataManager.CreateWaterTable();
                else return;
            }
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            DataTable dt = DataManager.GetWaterData(_dtpStart.Value.ToString("yyyy-MM-dd"), _dtpEnd.Value.ToString("yyyy-MM-dd"));
            _dgv.DataSource = dt;
            if (_dgv.Columns.Contains("Id")) _dgv.Columns["Id"].ReadOnly = true;
        }

        private void BtnSaveManual_Click(object sender, EventArgs e)
        {
            _dgv.EndEdit(); 
            DataTable dt = (DataTable)_dgv.DataSource;

            if (dt == null) return;

            int count = 0;
            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row.RowState == DataRowState.Deleted) continue;
                    
                    DataManager.UpsertWaterRecord(row);
                    count++;
                }
                MessageBox.Show(string.Format("成功處理 {0} 筆資料的新增/複寫作業。", count), "存檔完成");
                RefreshGrid(); 
            }
            catch (Exception ex)
            {
                MessageBox.Show("手動存檔失敗：" + ex.Message);
            }
        }

        // 🟢 執行刪除邏輯
        private void BtnDeleteRow_Click(object sender, EventArgs e)
        {
            // 檢查是否有點選資料列
            if (_dgv.CurrentRow == null)
            {
                MessageBox.Show("請先點選您要刪除的資料行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 判斷是否為尚未存檔的空白新列
            if (_dgv.CurrentRow.IsNewRow)
            {
                MessageBox.Show("無法刪除尚未存檔的空白行。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var idCell = _dgv.CurrentRow.Cells["Id"].Value;

            // 再次確認該列是否有對應的資料庫 ID
            if (idCell == null || idCell == DBNull.Value)
            {
                _dgv.Rows.Remove(_dgv.CurrentRow); // 僅存在於畫面的未儲存資料，直接移除
                return;
            }

            int id = Convert.ToInt32(idCell);

            // 跳出再次確認視窗
            var confirmResult = MessageBox.Show(
                "確定要刪除這筆紀錄嗎？\n刪除後將無法復原！", 
                "刪除確認", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Warning);

            if (confirmResult == DialogResult.Yes)
            {
                DataManager.DeleteWaterRecord(id);
                MessageBox.Show("資料已成功刪除！", "刪除完成");
                RefreshGrid(); // 重新載入資料表
            }
        }

        private void BtnAddCol_Click(object sender, EventArgs e)
        {
            string colName = _txtNewColName.Text.Trim();
            if (string.IsNullOrEmpty(colName)) return;
            DataManager.AddColumnToWaterTable(colName);
            MessageBox.Show("欄位新增成功，請重新讀取。");
            _txtNewColName.Clear();
        }
    }
}
