/// FILE: Safety_System/CoreTable/App_CoreTable.Events.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public partial class App_CoreTable
    {
        private bool _isBulkUpdating = false;

        private void BindEvents()
        {
            // Grid 視覺與互動事件
            _dgv.CellFormatting += new DataGridViewCellFormattingEventHandler(Dgv_CellFormatting);
            _dgv.CellClick += new DataGridViewCellEventHandler(Dgv_CellClick);
            _dgv.CellMouseClick += new DataGridViewCellMouseEventHandler(Dgv_CellMouseClick); 
            _dgv.KeyDown += new KeyEventHandler(Dgv_KeyDown);
            _dgv.KeyPress += new KeyPressEventHandler(Dgv_KeyPress);
            _dgv.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(Dgv_EditingControlShowing);
            
            _dgv.DataError += delegate(object s, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            
            _dgv.CurrentCellDirtyStateChanged += new EventHandler(Dgv_CurrentCellDirtyStateChanged);
            _dgv.CellValueChanged += new DataGridViewCellEventHandler(Dgv_CellValueChanged);
            
            _dgv.Sorted += new EventHandler(Dgv_Sorted);

            _dgv.ColumnWidthChanged += new DataGridViewColumnEventHandler(Dgv_ColumnWidthChanged);
            _dgv.ColumnDisplayIndexChanged += new DataGridViewColumnEventHandler(Dgv_ColumnDisplayIndexChanged);

            _dgv.RowValidated += new DataGridViewCellEventHandler(Dgv_RowValidated);

            _btnSave.Click += new EventHandler(BtnSave_Click);

            _btnRead.Click += new EventHandler(async delegate(object s, EventArgs e) {
                _isFirstLoad = false;
                _currentSearchMode = SearchMode.DateRange;
                await ReloadCurrentDataAsync();
            });

            Button bLimitRead = _btnRead.Tag as Button;
            if (bLimitRead != null) {
                bLimitRead.Click += new EventHandler(async delegate(object s, EventArgs e) {
                    _isFirstLoad = false;
                    _currentSearchMode = SearchMode.Limit;
                    
                    int limit = 100; // 預設改為 100
                    if (int.TryParse(_txtLatestCount.Text, out limit)) _currentLimit = limit;
                    
                    await ReloadCurrentDataAsync();
                });
            }

            _btnAdvancedSearch.Click += new EventHandler(async delegate(object s, EventArgs e) {
                _isFirstLoad = false;
                _currentSearchMode = SearchMode.Advanced;
                await ReloadCurrentDataAsync();
            });

            _btnExport.Click += new EventHandler(BtnExport_Click);
            _btnImport.Click += new EventHandler(BtnImportExcel_Click);
            _btnExportPdf.Click += new EventHandler(BtnExportPdf_Click); // 綁定匯出 PDF 彈窗事件
            _btnColSettings.Click += new EventHandler(BtnColSettings_Click);

            Button btnDelRow = _ctxMenu.Tag as Button;
            if (btnDelRow != null) {
                btnDelRow.Click += delegate(object s, EventArgs e) { ExecuteDeleteRow(); };
            }

            foreach (ToolStripItem item in _ctxMenu.Items)
            {
                if (item.Tag != null) {
                    string tagStr = item.Tag.ToString();
                    if (tagStr == "Save") item.Click += new EventHandler(BtnSave_Click);
                    if (tagStr == "DeleteRow") item.Click += delegate(object s, EventArgs e) { ExecuteDeleteRow(); };
                    if (tagStr == "OpenUrl") item.Click += new EventHandler(BtnOpenUrl_Click); // 開啟網址事件
                    if (tagStr == "ColSettings") item.Click += new EventHandler(BtnColSettings_Click);
                    if (tagStr == "Import") item.Click += new EventHandler(BtnImportExcel_Click);
                    if (tagStr == "Export") item.Click += new EventHandler(BtnExport_Click);
                    if (tagStr == "Pdf") item.Click += new EventHandler(BtnExportPdf_Click);
                }
            }

            if (_btnRtfToExcel != null) {
                _btnRtfToExcel.Click += new EventHandler(BtnRtfToExcel_Click);
            }
        }

        // 🟢 實作：開啟網址功能
        private void BtnOpenUrl_Click(object sender, EventArgs e)
        {
            if (_rightClickedColIndex >= 0 && _dgv.CurrentCell != null)
            {
                int r = _dgv.CurrentCell.RowIndex;
                int c = _rightClickedColIndex;
                
                object val = _dgv.Rows[r].Cells[c].Value;
                if (val != null)
                {
                    string url = val.ToString().Trim();
                    if (string.IsNullOrWhiteSpace(url)) return;

                    if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && 
                        !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    {
                        url = "http://" + url;
                    }

                    string browser = DataManager.GetSysSetting("DefaultBrowser", "chrome.exe");

                    try { System.Diagnostics.Process.Start(browser, url); } 
                    catch {
                        try { System.Diagnostics.Process.Start(url); } 
                        catch (Exception ex) {
                            MessageBox.Show($"無法開啟網址！\n請確認網址格式是否正確：\n\n{url}\n\n錯誤訊息：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void Dgv_Sorted(object sender, EventArgs e)
        {
            if (!_isApplyingWidths && _dgv.Rows.Count > 0)
            {
                try { _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders); } catch { }
            }
        }

        private void Dgv_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (_isBulkUpdating || _dgv.DataSource == null) return;
            
            try 
            {
                if (e.RowIndex >= 0 && e.RowIndex < _dgv.Rows.Count && !_dgv.Rows[e.RowIndex].IsNewRow) 
                {
                    DataRowView drv = _dgv.Rows[e.RowIndex].DataBoundItem as DataRowView;
                    if (drv != null) 
                    {
                        DataRow row = drv.Row;
                        if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified) 
                        {
                            Task.Run(() => 
                            {
                                try 
                                {
                                    DataManager.UpsertRecord(_dbName, _tableName, row);
                                    if (_dgv.InvokeRequired) {
                                        _dgv.Invoke(new Action(() => row.AcceptChanges()));
                                    } else {
                                        row.AcceptChanges();
                                    }
                                } catch { }
                            });
                        }
                    }
                }
            } 
            catch { }
        }

        private async void BtnSave_Click(object sender, EventArgs e) 
        {
            try {
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;
                _btnSave.Focus();

                if (_dgv.IsCurrentCellInEditMode) _dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
                _dgv.EndEdit(); 
                
                if (_dgv.BindingContext != null && _dgv.DataSource != null) _dgv.BindingContext[_dgv.DataSource].EndCurrentEdit();

                int scrollRow = _dgv.FirstDisplayedScrollingRowIndex;
                int scrollCol = _dgv.FirstDisplayedScrollingColumnIndex;
                int curRow = _dgv.CurrentCell?.RowIndex ?? -1;
                int curCol = _dgv.CurrentCell?.ColumnIndex ?? -1;

                SaveColumnOrder(); 
                SetUIState(false, "資料庫寫入與檔案同步中，請稍候...", Color.Orange);
                
                DataTable dtSource = (DataTable)_dgv.DataSource;
                DataTable dtChanges = dtSource.GetChanges(); 

                if (dtChanges == null || dtChanges.Rows.Count == 0)
                {
                    SetUIState(true, "無異動資料需要儲存。", Color.Green); 
                    return; 
                }

                bool success = false;
                
                using (ProgressForm progForm = new ProgressForm("儲存變更的數據中..."))
                {
                    await progForm.ExecuteAsync(async delegate(IProgress<int> progInt, IProgress<string> progStr) 
                    { 
                        progStr.Report("正在格式化資料與同步附件路徑...");
                        EnforceDateFormats(dtChanges); 
                        SyncAttachmentPaths(dtChanges);

                        progStr.Report("正在執行模組預處理...");
                        if (await _logic.OnBeforeSaveAsync(_dbName, _tableName, dtChanges, progInt, progStr)) 
                        {
                            success = DataManager.BulkSaveTable(_dbName, _tableName, dtChanges, progInt, progStr); 
                            
                            if (success) {
                                progStr.Report("正在執行儲存後處理...");
                                await _logic.OnAfterSaveAsync(_dbName, _tableName, dtChanges);
                            }
                        }
                    });
                }
                
                if (success) { 
                    dtSource.AcceptChanges(); 

                    try { _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders); } catch { }

                    try {
                        if (scrollRow >= 0 && scrollRow < _dgv.Rows.Count) _dgv.FirstDisplayedScrollingRowIndex = scrollRow;
                        if (scrollCol >= 0 && scrollCol < _dgv.Columns.Count) _dgv.FirstDisplayedScrollingColumnIndex = scrollCol;
                        if (curRow >= 0 && curRow < _dgv.Rows.Count && curCol >= 0 && curCol < _dgv.Columns.Count) {
                            _dgv.CurrentCell = _dgv[curCol, curRow];
                        }
                    } catch { }

                    SetUIState(true, "資料儲存成功！", Color.Green); 
                    MessageBox.Show("儲存完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                } else { 
                    SetUIState(true, "資料儲存失敗", Color.Red); 
                }
            } catch (Exception ex) { 
                SetUIState(true, "儲存異常", Color.Red); 
                MessageBox.Show("儲存異常：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            } finally { 
                if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default; 
            }
        }

        private void ExecuteDeleteRow()
        {
            List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewCell cell in _dgv.SelectedCells) {
                DataGridViewRow r = cell.OwningRow;
                if (!r.IsNewRow && r.Cells["Id"].Value != DBNull.Value) {
                    if (!selectedRows.Contains(r)) {
                        selectedRows.Add(r);
                    }
                }
            }

            if (selectedRows.Count > 0 && MessageBox.Show($"確定要刪除選取的 {selectedRows.Count} 筆資料嗎？\n(刪除後系統將自動重算受影響之日期差值)", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                if (AuthManager.VerifyUser()) {
                    
                    if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.WaitCursor;

                    try
                    {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        foreach (DataGridViewRow r in selectedRows) {
                            if (_dgv.Columns.Contains("附件檔案")) {
                                object val = r.Cells["附件檔案"].Value;
                                if (val != null) {
                                    string relPathStr = val.ToString();
                                    if (!string.IsNullOrEmpty(relPathStr)) {
                                        string[] paths = relPathStr.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                        foreach (string p in paths) {
                                            if (!DataManager.IsAttachmentInUse(_dbName, _tableName, p))
                                            {
                                                try {
                                                    string absPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, p);
                                                    if (File.Exists(absPath)) File.Delete(absPath);
                                                } catch { }
                                            }
                                        }
                                    }
                                }
                            }

                            int id = Convert.ToInt32(r.Cells["Id"].Value);
                            DataManager.DeleteRecord(_dbName, _tableName, id);
                            
                            for (int i = 0; i < dt.Rows.Count; i++) {
                                DataRow dr = dt.Rows[i];
                                if (dr.RowState != DataRowState.Deleted && Convert.ToInt32(dr["Id"]) == id) {
                                    dr.Delete();
                                    break;
                                }
                            }
                        }
                        
                        dt.AcceptChanges(); 

                        try { _dgv.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders); } catch { }

                        if (_calcHelper != null) {
                            _calcHelper.BeginBulkUpdate();
                            _calcHelper.RecalculateTable(dt);
                            _calcHelper.EndBulkUpdate();
                        }

                        Task.Run(() => 
                        {
                            try 
                            {
                                foreach (DataRow row in dt.Rows) 
                                {
                                    if (row.RowState == DataRowState.Modified) 
                                    {
                                        DataManager.UpsertRecord(_dbName, _tableName, row);
                                        if (_dgv.InvokeRequired) {
                                            _dgv.Invoke(new Action(() => row.AcceptChanges()));
                                        } else {
                                            row.AcceptChanges();
                                        }
                                    }
                                }
                            } catch { }
                        });

                        MessageBox.Show("刪除成功！系統已將資料斷點無縫重算完畢。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    finally
                    {
                        if (Form.ActiveForm != null) Form.ActiveForm.Cursor = Cursors.Default;
                    }
                }
            }
        }

        private void BtnExport_Click(object sender, EventArgs e) 
        {
            DataTable dt = (DataTable)_dgv.DataSource;

            Dictionary<string, string[]> dropdownData = new Dictionary<string, string[]>();
            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (col is DataGridViewComboBoxColumn) {
                    DataGridViewComboBoxColumn cboCol = (DataGridViewComboBoxColumn)col;
                    List<string> itemsList = new List<string>();
                    foreach (object obj in cboCol.Items) {
                        itemsList.Add(obj.ToString());
                    }
                    if (itemsList.Count > 0) {
                        dropdownData[col.Name] = itemsList.ToArray();
                    }
                }
            }

            Dictionary<string, float> exportWidths = new Dictionary<string, float>();
            foreach (KeyValuePair<string, int> kvp in _columnWidths) {
                exportWidths[kvp.Key] = (float)kvp.Value;
            }

            ExcelHelper.ExportToExcelOrCsv(dt, _chineseTitle, exportWidths, dropdownData);
        }

        private async void BtnImportExcel_Click(object sender, EventArgs e) 
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", Title = "請選擇要匯入的 Excel 檔案" }) 
            {
                if (ofd.ShowDialog() == DialogResult.OK) 
                {
                    DataTable boundDt = (DataTable)_dgv.DataSource; 
                    DataTable workingDt = boundDt.Copy(); 
                    DataTable templateDt = boundDt.Clone(); 

                    _isBulkUpdating = true; 
                    
                    using (ProgressForm progForm = new ProgressForm("匯入與運算中..."))
                    {
                        await progForm.ExecuteAsync(async delegate(IProgress<int> progInt, IProgress<string> progStr) 
                        {
                            DataTable importedDt = await ExcelHelper.ImportToDataTableAsync(ofd.FileName, templateDt, progInt, progStr);

                            if (importedDt != null && importedDt.Rows.Count > 0) {
                                progStr.Report("正在將資料合併至系統...");
                                progInt.Report(0);
                                
                                foreach (DataRow row in importedDt.Rows) workingDt.ImportRow(row);

                                if (_calcHelper != null) {
                                    _calcHelper.BeginBulkUpdate(); 
                                    _calcHelper.RecalculateTable(workingDt, progInt, progStr); 
                                    _calcHelper.EndBulkUpdate(); 
                                }
                                
                                progStr.Report("正在格式化資料...");
                                EnforceDateFormats(workingDt);
                            }
                        });
                    }

                    UnfreezeAllColumns();
                    _isApplyingWidths = true;
                    _dgv.SuspendLayout();

                    PreFillComboBoxItems(workingDt); 
                    _dgv.DataSource = workingDt; 
                    ApplyGridStyles(); 
                    RestoreColumnOrder();
                    ApplyFreezeState();

                    _dgv.ResumeLayout(true);
                    _isApplyingWidths = false;
                    _isBulkUpdating = false; 

                    SetUIState(true, $"Excel 匯入完成！新增資料後總筆數：{workingDt.Rows.Count}", Color.Green);
                    
                    if (MessageBox.Show("Excel 匯入運算完成！\n是否要【立即寫入】資料庫？\n(選擇「否」則可先檢查畫面，後續再手動點擊「儲存」按鈕)", "確認寫入", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        _btnSave.PerformClick(); 
                    }
                }
            }
        }

        // 🟢 實作：匯出 PDF 前，跳出選擇紙張大小與方向視窗
        private void BtnExportPdf_Click(object sender, EventArgs e) 
        {
            if (_dgv.Rows.Count <= 1 && _dgv.AllowUserToAddRows) 
            { 
                MessageBox.Show("目前沒有資料可供導出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                return; 
            }

            using (Form f = new Form())
            {
                f.Width = 350;
                f.Height = 250;
                f.Text = "PDF 輸出設定";
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;
                f.BackColor = Color.White;

                Label lbl1 = new Label { Text = "請選擇紙張大小：", Location = new Point(30, 30), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                ComboBox cboSize = new ComboBox { Location = new Point(170, 27), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                cboSize.Items.AddRange(new string[] { "A4", "A3" });
                cboSize.SelectedIndex = 0;

                Label lbl2 = new Label { Text = "請選擇紙張方向：", Location = new Point(30, 80), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                ComboBox cboLayout = new ComboBox { Location = new Point(170, 77), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                cboLayout.Items.AddRange(new string[] { "直式", "橫式" });
                cboLayout.SelectedIndex = 1; // 預設橫式

                Button btnOk = new Button { Text = "確認導出", Location = new Point(110, 140), Size = new Size(120, 40), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnOk.Click += delegate(object senderObj, EventArgs ev) {
                    f.DialogResult = DialogResult.OK;
                };

                f.Controls.Add(lbl1); f.Controls.Add(cboSize);
                f.Controls.Add(lbl2); f.Controls.Add(cboLayout);
                f.Controls.Add(btnOk);

                if (f.ShowDialog(Form.ActiveForm) == DialogResult.OK)
                {
                    bool isA3 = cboSize.SelectedItem.ToString() == "A3";
                    bool isLandscape = cboLayout.SelectedItem.ToString() == "橫式";
                    
                    // 呼叫 PdfHelper 執行匯出，傳遞參數
                    PdfHelper.ExportDataGridViewToPdf(_dgv, _chineseTitle, _chineseTitle, isA3, isLandscape);
                }
            }
        }

        private void BtnColSettings_Click(object sender, EventArgs e) 
        {
            if (_dgv.Columns.Count == 0) return;
            using (Form f = new Form { Text = "👁️ 欄位顯示設定", Size = new Size(350, 500), StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false }) 
            {
                Label lblTop = new Label { Text = "請勾選欲顯示在表格中的欄位：", Dock = DockStyle.Top, Padding = new Padding(10), Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), ForeColor = Color.SteelBlue }; 
                f.Controls.Add(lblTop);
                
                CheckedListBox clbCols = new CheckedListBox { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), CheckOnClick = true, BorderStyle = BorderStyle.None, Padding = new Padding(10) };
                
                foreach (DataGridViewColumn col in _dgv.Columns) { 
                    if (col.Name == "Id") continue; 
                    bool isChecked = true;
                    if (_columnVisibility.ContainsKey(col.Name)) isChecked = _columnVisibility[col.Name];
                    clbCols.Items.Add(col.Name, isChecked); 
                }

                clbCols.ItemCheck += (s, ev) => {
                    string colName = clbCols.Items[ev.Index].ToString();
                    if ((colName == "最後修改人" || colName == "修改時間") && ev.NewValue == CheckState.Checked) {
                        if (!AuthManager.VerifyLv3Only("顯示修改紀錄需要系統管理者權限\n請輸入【Lv3系統管理者】密碼：")) {
                            ev.NewValue = CheckState.Unchecked;
                        }
                    }
                };
                
                f.Controls.Add(clbCols);
                
                Button btnSaveLocal = new Button { Text = "💾 儲存並套用設定", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
                btnSaveLocal.Click += delegate(object s, EventArgs ev) { 
                    for (int i = 0; i < clbCols.Items.Count; i++) { 
                        string colName = clbCols.Items[i].ToString(); 
                        bool isChecked = clbCols.GetItemChecked(i); 
                        _columnVisibility[colName] = isChecked; 
                        if (_dgv.Columns.Contains(colName)) _dgv.Columns[colName].Visible = isChecked; 
                    } 
                    SaveVisibilitySettings(); 
                    f.DialogResult = DialogResult.OK; 
                };
                
                f.Controls.Add(btnSaveLocal); 
                f.ShowDialog();
                ApplyFreezeState(); 
            }
        }

        private void BtnRtfToExcel_Click(object sender, EventArgs e) 
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "RTF 法規檔案 (*.rtf)|*.rtf", Title = "請選擇全國法規資料庫下載的 RTF 檔案" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel 檔案 (*.xlsx)|*.xlsx", FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + "_轉換.xlsx" }) {
                        if (sfd.ShowDialog() == DialogResult.OK) {
                            try { 
                                LawRtfToExcelConverter.Convert(ofd.FileName, sfd.FileName); 
                                MessageBox.Show("轉換成功！\n您現在可以點擊「匯入」將產生的檔案載入系統。", "轉換完成", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                            } catch (Exception ex) { 
                                MessageBox.Show("轉換失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                            }
                        }
                    }
                }
            }
        }

        private async void Dgv_KeyDown(object sender, KeyEventArgs e) 
        {
            if (e.Control && e.KeyCode == Keys.S) { 
                e.Handled = true; e.SuppressKeyPress = true; 
                if (_btnSave != null) _btnSave.PerformClick(); 
            }
            else if ((e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) && !_dgv.IsCurrentCellInEditMode) {
                bool hasCleared = false;
                foreach (DataGridViewCell cell in _dgv.SelectedCells) {
                    if (!cell.ReadOnly && cell.OwningColumn.Name != "Id" && cell.OwningColumn.Name != "附件檔案" && !cell.OwningRow.IsNewRow) {
                        cell.Value = DBNull.Value; hasCleared = true;
                    }
                }
                
                if (hasCleared) {
                    e.Handled = true;
                    if (_dgv.DataSource is DataTable) {
                        DataTable dt = (DataTable)_dgv.DataSource;
                        if (_calcHelper != null) {
                            _calcHelper.BeginBulkUpdate();
                            _calcHelper.RecalculateTable(dt);
                            _calcHelper.EndBulkUpdate();
                        }
                        EnforceDateFormats(dt);
                    }
                }
            }
            else if (e.Control && e.KeyCode == Keys.V) {
                string text = Clipboard.GetText(); 
                if (string.IsNullOrEmpty(text)) return;
                
                string[] splitChars = new string[] { "\r\n", "\r", "\n" };
                string[] lines = text.Split(splitChars, StringSplitOptions.RemoveEmptyEntries);
                
                int r = 0;
                int c = 0;
                if (_dgv.CurrentCell != null) {
                    r = _dgv.CurrentCell.RowIndex;
                    c = _dgv.CurrentCell.ColumnIndex;
                }
                
                DataTable boundDt = (DataTable)_dgv.DataSource;
                DataTable workingDt = boundDt.Copy();

                List<int> readOnlyCols = new List<int>();
                for (int i = 0; i < _dgv.Columns.Count; i++) {
                    if (_dgv.Columns[i].ReadOnly) readOnlyCols.Add(i);
                }

                _isBulkUpdating = true; 

                using (ProgressForm progForm = new ProgressForm("貼上資料與運算中...")) {
                    await progForm.ExecuteAsync(async delegate(IProgress<int> progInt, IProgress<string> progStr) {
                        progStr.Report("正在解析貼上的資料...");
                        
                        foreach (string line in lines) {
                            if (r >= workingDt.Rows.Count) workingDt.Rows.Add(workingDt.NewRow());
                            string[] cells = line.Split('\t');
                            for (int i = 0; i < cells.Length; i++) { 
                                if (c + i < workingDt.Columns.Count && !readOnlyCols.Contains(c + i)) { 
                                    workingDt.Rows[r][c + i] = cells[i].Trim().Trim('"'); 
                                } 
                            }
                            r++;
                        }

                        if (_calcHelper != null) {
                            _calcHelper.BeginBulkUpdate(); 
                            _calcHelper.RecalculateTable(workingDt, progInt, progStr); 
                            _calcHelper.EndBulkUpdate(); 
                        }
                        EnforceDateFormats(workingDt); 
                    });
                }

                UnfreezeAllColumns();
                _isApplyingWidths = true;
                _dgv.SuspendLayout();

                PreFillComboBoxItems(workingDt); 
                _dgv.DataSource = workingDt; 
                ApplyGridStyles(); 
                RestoreColumnOrder();
                ApplyFreezeState();

                _dgv.ResumeLayout(true);
                _isApplyingWidths = false;
                _isBulkUpdating = false; 
            }
        }

        private void Dgv_KeyPress(object sender, KeyPressEventArgs e) 
        {
            if (_dgv.CurrentCell != null && !_dgv.CurrentCell.ReadOnly && !_dgv.IsCurrentCellInEditMode) {
                if (char.IsLetterOrDigit(e.KeyChar) || char.IsPunctuation(e.KeyChar) || char.IsSymbol(e.KeyChar) || char.IsWhiteSpace(e.KeyChar)) {
                    _dgv.BeginEdit(true); 
                    if (_dgv.EditingControl is TextBox) { 
                        TextBox txt = (TextBox)_dgv.EditingControl;
                        txt.Text = e.KeyChar.ToString(); 
                        txt.SelectionStart = txt.Text.Length; 
                        e.Handled = true; 
                    }
                }
            }
        }

        private void Dgv_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e) 
        {
            e.Control.PreviewKeyDown -= new PreviewKeyDownEventHandler(EditingControl_PreviewKeyDown);
            e.Control.PreviewKeyDown += new PreviewKeyDownEventHandler(EditingControl_PreviewKeyDown);

            if (e.Control is ComboBox) {
                ComboBox cbo = (ComboBox)e.Control;
                cbo.DropDownStyle = ComboBoxStyle.DropDownList;
                if (_dgv.CurrentCell != null) {
                    string colName = _dgv.Columns[_dgv.CurrentCell.ColumnIndex].Name;
                    string parentColName = "";
                    
                    foreach (KeyValuePair<string, string[]> kvp in App_DropdownManager.DropdownCache) {
                        string[] parts = kvp.Key.Split('|');
                        if (parts.Length == 4 && parts[0] == _tableName && parts[1] == colName && !string.IsNullOrEmpty(parts[2])) {
                            parentColName = parts[2]; break;
                        }
                    }

                    string[] items = null;
                    if (!string.IsNullOrEmpty(parentColName) && _dgv.Columns.Contains(parentColName)) {
                        string parentVal = "";
                        object rawVal = _dgv.CurrentRow.Cells[parentColName].Value;
                        if (rawVal != null) parentVal = rawVal.ToString();

                        items = App_DropdownManager.GetOptions(_tableName, colName, parentColName, parentVal);
                        if (items == null || items.Length == 0) items = _logic.GetDependentDropdownList(_tableName, colName, parentVal); 
                    } else {
                        items = App_DropdownManager.GetAllOptionsForColumn(_tableName, colName);
                        if (items == null || items.Length <= 1) items = _logic.GetDropdownList(_tableName, colName);
                    }
                    
                    if (items != null) { 
                        object currentVal = _dgv.CurrentCell.Value; 
                        cbo.Items.Clear(); 
                        cbo.Items.AddRange(items); 
                        if (currentVal != null && cbo.Items.Contains(currentVal)) cbo.SelectedItem = currentVal; 
                    }
                }
            }
            else if (e.Control is TextBox) { 
                TextBox txt = (TextBox)e.Control;
                txt.Multiline = true; 
                txt.KeyDown -= new KeyEventHandler(TextBox_KeyDown); 
                txt.KeyDown += new KeyEventHandler(TextBox_KeyDown); 
            }
        }

        private void EditingControl_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e) {
            if (e.Control && e.KeyCode == Keys.S) { e.IsInputKey = true; _btnSave.PerformClick(); }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.Alt && e.KeyCode == Keys.Enter) { 
                if (sender is TextBox) { 
                    TextBox txt = (TextBox)sender;
                    int selectionStart = txt.SelectionStart; 
                    txt.Text = txt.Text.Insert(selectionStart, Environment.NewLine); 
                    txt.SelectionStart = selectionStart + Environment.NewLine.Length; 
                    e.Handled = true; 
                } 
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e) 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < _dgv.Rows.Count && !_dgv.Rows[e.RowIndex].IsNewRow) {
                if (_dgv.Columns[e.ColumnIndex].Name.Contains("附件檔案")) {
                    string currentVal = "";
                    object rawVal = _dgv[e.ColumnIndex, e.RowIndex].Value;
                    if (rawVal != null) currentVal = rawVal.ToString();

                    string rowDateStr = "";
                    object rawDateVal = _dgv[_dateColumnName, e.RowIndex].Value;
                    if (rawDateVal != null) rowDateStr = rawDateVal.ToString();

                    string targetFolder = GetExpectedFolderName(rowDateStr);
                    
                    using (var frm = new AttachmentManagerUI(currentVal, _dbName, _tableName, targetFolder, delegate(string path) { DeletePhysicalFile(path, e.RowIndex); })) {
                        if (frm.ShowDialog() == DialogResult.OK) { 
                            _dgv[e.ColumnIndex, e.RowIndex].Value = frm.FinalPathsString; 
                            _dgv.EndEdit(); 
                        }
                    }
                } else { 
                    _logic.OnCellClick(_dgv, e); 
                }
            }
        }

        private void Dgv_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) {
                if (e.ColumnIndex >= 0) {
                    _rightClickedColIndex = e.ColumnIndex;
                    if (e.RowIndex >= 0 && e.RowIndex < _dgv.Rows.Count) {
                        if (!_dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected) {
                            _dgv.ClearSelection();
                            _dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                            _dgv.CurrentCell = _dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        }
                    }
                    _ctxMenu.Show(Cursor.Position);
                }
            }
        }

        private void Dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || _isCascading) return;
            string colName = _dgv.Columns[e.ColumnIndex].Name;

            if (!_isApplyingWidths && !_isFirstLoad) {
                try { _dgv.AutoResizeRow(e.RowIndex, DataGridViewAutoSizeRowMode.AllCellsExceptHeader); } catch { }
            }

            _isCascading = true;
            try {
                foreach (KeyValuePair<string, string[]> kvp in App_DropdownManager.DropdownCache) {
                    string[] parts = kvp.Key.Split('|');
                    if (parts.Length == 4 && parts[0] == _tableName && parts[2] == colName) {
                        string childColName = parts[1];
                        if (_dgv.Columns.Contains(childColName)) {
                            if (e.RowIndex < _dgv.Rows.Count) _dgv.Rows[e.RowIndex].Cells[childColName].Value = DBNull.Value;
                        }
                    }
                }
            } catch { } finally { _isCascading = false; }
        }

        private void Dgv_CurrentCellDirtyStateChanged(object sender, EventArgs e) {
            if (_dgv.IsCurrentCellDirty && _dgv.CurrentCell is DataGridViewComboBoxCell) { _dgv.CommitEdit(DataGridViewDataErrorContexts.Commit); }
        }

        private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) {
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                if (colName.Contains("附件檔案") && e.Value != null) {
                    string pathStr = e.Value.ToString();
                    if (!string.IsNullOrEmpty(pathStr)) {
                        string[] parts = pathStr.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 1) { e.Value = $"📁 [共 {parts.Length} 個檔案]"; } 
                        else { e.Value = Path.GetFileName(parts[0]); }
                        e.FormattingApplied = true;
                    }
                }
            }
        }

        private void Dgv_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e) {
            if (_isFirstLoad || _isApplyingWidths) return;
            if (e.Column != null && e.Column.Visible && e.Column.Width > 0 && e.Column.AutoSizeMode == DataGridViewAutoSizeColumnMode.None) { 
                _columnWidths[e.Column.Name] = e.Column.Width; 
                DataManager.SaveGridConfig(_dbName, _tableName, "Width", e.Column.Name, e.Column.Width.ToString());
            }
        }

        private void Dgv_ColumnDisplayIndexChanged(object sender, DataGridViewColumnEventArgs e) {
            if (_isFirstLoad || _isApplyingWidths) return;
            SaveColumnOrder();
        }
    }
}
