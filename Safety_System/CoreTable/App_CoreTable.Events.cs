using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq; 
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public partial class App_CoreTable
    {
        private bool _isBulkUpdating = false;

        // 智慧防呆緩衝器：紀錄每一列資料的延遲儲存任務
        private Dictionary<DataRow, CancellationTokenSource> _rowSaveTimers = new Dictionary<DataRow, CancellationTokenSource>();

        private void BindEvents()
        {
            _dgv.CellFormatting += new DataGridViewCellFormattingEventHandler(Dgv_CellFormatting);
            _dgv.CellClick += new DataGridViewCellEventHandler(Dgv_CellClick);
            _dgv.CellDoubleClick += new DataGridViewCellEventHandler(Dgv_CellDoubleClick);
            _dgv.CellMouseClick += new DataGridViewCellMouseEventHandler(Dgv_CellMouseClick); 
            _dgv.KeyDown += new KeyEventHandler(Dgv_KeyDown);
            _dgv.KeyPress += new KeyPressEventHandler(Dgv_KeyPress);
            _dgv.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(Dgv_EditingControlShowing);
            
            // 核心攔截：儲存格自訂繪製 (純圖示處理)
            _dgv.CellPainting += new DataGridViewCellPaintingEventHandler(Dgv_CellPainting);
            
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
                    
                    int limit = 100; 
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
            _btnExportPdf.Click += new EventHandler(BtnExportPdf_Click); 
            _btnColSettings.Click += new EventHandler(BtnColSettings_Click);

            Button btnDelRow = _ctxMenu.Tag as Button;
            if (btnDelRow != null) {
                btnDelRow.Click += delegate(object s, EventArgs e) { ExecuteDeleteRow(); };
            }

            _cboSearchColumn.SelectedIndexChanged += CboSearchColumn_SelectedIndexChanged;

            foreach (ToolStripItem item in _ctxMenu.Items)
            {
                if (item.Tag != null) {
                    string tagStr = item.Tag.ToString();
                    if (tagStr == "Save") item.Click += new EventHandler(BtnSave_Click);
                    if (tagStr == "DeleteRow") item.Click += delegate(object s, EventArgs e) { ExecuteDeleteRow(); };
                    if (tagStr == "OpenUrl") item.Click += new EventHandler(BtnOpenUrl_Click); 
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

        private void CboSearchColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isCascading) return; 

            if (_cboSearchColumn.SelectedItem == null || string.IsNullOrWhiteSpace(_cboSearchColumn.SelectedItem.ToString()))
            {
                _isSearchDropdownMode = false;
                _cboSearchKeyword.Visible = false;
                _txtSearchKeyword.Visible = true;
                _txtSearchKeyword.Clear();
                return;
            }

            string selectedCol = _cboSearchColumn.SelectedItem.ToString();
            string multiKey = $"{_tableName}|{selectedCol}";
            string[] items = null;

            if (App_DropdownManager.MultiSelectCache.ContainsKey(multiKey)) {
                items = App_DropdownManager.MultiSelectCache[multiKey].Select(x => x.Text).ToArray();
            } 
            else {
                string[] dbItems = App_DropdownManager.GetAllOptionsForColumn(_tableName, selectedCol);
                if (dbItems != null && dbItems.Length > 1) {
                    items = dbItems;
                } else {
                    string[] logicItems = _logic.GetDropdownList(_tableName, selectedCol);
                    if (logicItems != null && logicItems.Length > 0) items = logicItems;
                }
            }

            if (items != null && items.Length > 0)
            {
                _isSearchDropdownMode = true;
                _txtSearchKeyword.Visible = false;
                _cboSearchKeyword.Visible = true;
                
                _cboSearchKeyword.Items.Clear();
                _cboSearchKeyword.Items.Add(""); 
                
                foreach (string item in items) {
                    string cleanItem = item.Trim();
                    if (!string.IsNullOrEmpty(cleanItem) && !_cboSearchKeyword.Items.Contains(cleanItem)) {
                        _cboSearchKeyword.Items.Add(cleanItem);
                    }
                }
                
                _cboSearchKeyword.SelectedIndex = 0;
            }
            else
            {
                _isSearchDropdownMode = false;
                _cboSearchKeyword.Visible = false;
                _txtSearchKeyword.Visible = true;
                _txtSearchKeyword.Clear();
            }
        }

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

        private void Dgv_Sorted(object sender, EventArgs e) { }

        private async void Dgv_RowValidated(object sender, DataGridViewCellEventArgs e)
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
                            bool isNewData = (row.RowState == DataRowState.Added) || 
                                             (row.Table.Columns.Contains("Id") && (row["Id"] == DBNull.Value || Convert.ToInt32(row["Id"]) <= 0));

                            int delayMs = isNewData ? 10000 : 500;

                            if (_rowSaveTimers.ContainsKey(row)) 
                            {
                                _rowSaveTimers[row].Cancel();
                                _rowSaveTimers[row].Dispose();
                            }

                            CancellationTokenSource cts = new CancellationTokenSource();
                            _rowSaveTimers[row] = cts;

                            try 
                            {
                                await Task.Delay(delayMs, cts.Token);
                            } 
                            catch (TaskCanceledException) 
                            {
                                return; 
                            }

                            _rowSaveTimers.Remove(row);

                            await Task.Run(() => 
                            {
                                try 
                                {
                                    long newId = DataManager.UpsertRecord(_dbName, _tableName, row);
                                    
                                    if (_dgv.InvokeRequired) 
                                    {
                                        _dgv.Invoke(new Action(() => {
                                            if (newId > 0 && row.Table.Columns.Contains("Id")) {
                                                bool originalReadOnly = row.Table.Columns["Id"].ReadOnly;
                                                row.Table.Columns["Id"].ReadOnly = false;
                                                row["Id"] = newId;
                                                row.Table.Columns["Id"].ReadOnly = originalReadOnly;
                                            }
                                            row.AcceptChanges();
                                        }));
                                    } 
                                    else 
                                    {
                                        if (newId > 0 && row.Table.Columns.Contains("Id")) {
                                            bool originalReadOnly = row.Table.Columns["Id"].ReadOnly;
                                            row.Table.Columns["Id"].ReadOnly = false;
                                            row["Id"] = newId;
                                            row.Table.Columns["Id"].ReadOnly = originalReadOnly;
                                        }
                                        row.AcceptChanges();
                                    }
                                } 
                                catch { }
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

                foreach (var cts in _rowSaveTimers.Values) {
                    try { cts.Cancel(); cts.Dispose(); } catch { }
                }
                _rowSaveTimers.Clear();

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
                                                    string absPath = Path.Combine(DataManager.AttachmentBasePath, p.Replace("附件/", ""));
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
                if (col.Name != "Id" && col.Name != "附件檔案" && col.Name != "備註") {
                    string[] dbItems = App_DropdownManager.GetAllOptionsForColumn(_tableName, col.Name);
                    if (dbItems != null && dbItems.Length > 1) {
                        dropdownData[col.Name] = dbItems;
                    }
                }
            }

            Dictionary<string, float> exportWidths = new Dictionary<string, float>();
            foreach (KeyValuePair<string, int> kvp in _columnWidths) {
                exportWidths[kvp.Key] = (float)kvp.Value;
            }

            var formulas = DataManager.GetTableFormulas(_dbName, _tableName);
            HashSet<string> formulaCols = new HashSet<string>();
            if (formulas != null) {
                foreach (var f in formulas) {
                    formulaCols.Add(f.TargetCol);
                }
            }

            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (col.Name == "星期" || col.Name.EndsWith("日統計") || col.Name.EndsWith("月統計") || col.Name.EndsWith("年統計")) {
                    formulaCols.Add(col.Name);
                }
            }

            DataTable exportDt = new DataTable();
            List<DataGridViewColumn> visCols = new List<DataGridViewColumn>();

            foreach (DataGridViewColumn col in _dgv.Columns) {
                if (col.Visible && col.Name != "Id") {
                    visCols.Add(col);
                    exportDt.Columns.Add(col.HeaderText.Replace("\n", ""));
                }
            }

            foreach (DataGridViewRow row in _dgv.Rows) {
                if (row.IsNewRow) continue;
                DataRow dRow = exportDt.NewRow();
                for (int i = 0; i < visCols.Count; i++) {
                    var cellVal = row.Cells[visCols[i].Index].Value;
                    dRow[i] = cellVal != null ? cellVal.ToString() : "";
                }
                exportDt.Rows.Add(dRow);
            }

            ExcelHelper.ExportToExcelOrCsv(exportDt, _chineseTitle, exportWidths, dropdownData, formulaCols);
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
                                
                                foreach (DataRow importedRow in importedDt.Rows) 
                                {
                                    DataRow newRow = workingDt.NewRow();
                                    foreach (DataColumn col in importedDt.Columns) 
                                    {
                                        if (workingDt.Columns.Contains(col.ColumnName)) 
                                        {
                                            newRow[col.ColumnName] = importedRow[col.ColumnName];
                                        }
                                    }
                                    workingDt.Rows.Add(newRow);
                                }

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

                    _dgv.Visible = false;
                    UnfreezeAllColumns();
                    _isApplyingWidths = true;
                    _dgv.SuspendLayout();

                    _dgv.DataSource = workingDt; 
                    ApplyGridStyles(); 
                    RestoreColumnOrder();
                    ApplyFreezeState();

                    _dgv.ResumeLayout(true);
                    _dgv.Visible = true;
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

        private void BtnExportPdf_Click(object sender, EventArgs e) 
        {
            if (_dgv.Rows.Count <= 1 && _dgv.AllowUserToAddRows) 
            { 
                MessageBox.Show("目前沒有資料可供導出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                return; 
            }

            using (Form f = new Form())
            {
                f.Width = 420; 
                f.Height = 250;
                f.Text = "PDF 輸出設定";
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;
                f.BackColor = Color.White;

                Label lbl1 = new Label { Text = "請選擇紙張大小：", Location = new Point(30, 30), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                ComboBox cboSize = new ComboBox { Location = new Point(220, 27), Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                cboSize.Items.AddRange(new string[] { "A4", "A3" });
                cboSize.SelectedIndex = 0;

                Label lbl2 = new Label { Text = "請選擇紙張方向：", Location = new Point(30, 80), AutoSize = true, Font = new Font("Microsoft JhengHei UI", 12F) };
                ComboBox cboLayout = new ComboBox { Location = new Point(220, 77), Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft JhengHei UI", 12F) };
                cboLayout.Items.AddRange(new string[] { "直式", "橫式" });
                cboLayout.SelectedIndex = 1; 

                Button btnOk = new Button { Text = "確認導出", Location = new Point(140, 140), Size = new Size(120, 40), BackColor = Color.IndianRed, ForeColor = Color.White, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), Cursor = Cursors.Hand };
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
                    
                    PdfHelper.ExportDataGridViewToPdf(_dgv, _chineseTitle, _tableName, isA3, isLandscape);
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

        // ========================================================================
        // 🟢 核心重構：智慧欄位對齊貼上 (Smart Column Alignment Paste)
        // ========================================================================
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
            // 🟢 Ctrl+V 貼上邏輯優化：解決資料錯位與重複問題
            else if (e.Control && e.KeyCode == Keys.V) {
                string text = Clipboard.GetText(); 
                if (string.IsNullOrEmpty(text)) return;
                
                string[] splitChars = new string[] { "\r\n", "\r", "\n" };
                string[] lines = text.Split(splitChars, StringSplitOptions.RemoveEmptyEntries);
                
                if (_dgv.CurrentCell == null && _dgv.SelectedRows.Count == 0) return;

                // 1. 取得畫面上「真正可見」的欄位清單，並且依照使用者的拖曳排序 (DisplayIndex)
                var visibleCols = _dgv.Columns.Cast<DataGridViewColumn>()
                    .Where(col => col.Visible)
                    .OrderBy(col => col.DisplayIndex)
                    .ToList();

                if (visibleCols.Count == 0) return;

                int startRowIdx = 0;
                int visibleStartColIndex = 0;

                // 🟢 智慧對齊判斷：
                // 如果使用者是點擊「最左側把手」選取整列 (SelectedRows.Count > 0)，
                // 則強制從第一個可見欄位 (Index 0) 開始貼上，避免錯位！
                if (_dgv.SelectedRows.Count > 0)
                {
                    // 找出選取範圍中最上面的一列
                    startRowIdx = _dgv.SelectedRows.Cast<DataGridViewRow>().Min(r => r.Index);
                    visibleStartColIndex = 0; 
                }
                else
                {
                    // 單一或局部儲存格選取模式，尊重當下 CurrentCell 的位置
                    startRowIdx = _dgv.CurrentCell.RowIndex;
                    int startColIdx = _dgv.CurrentCell.ColumnIndex;
                    visibleStartColIndex = visibleCols.FindIndex(c => c.Index == startColIdx);
                    if (visibleStartColIndex == -1) return;
                }

                DataTable boundDt = (DataTable)_dgv.DataSource;
                DataTable workingDt = boundDt.Copy();

                _isBulkUpdating = true; 

                using (ProgressForm progForm = new ProgressForm("貼上資料與運算中...")) {
                    await progForm.ExecuteAsync(async delegate(IProgress<int> progInt, IProgress<string> progStr) {
                        progStr.Report("正在解析貼上的資料 (智慧對齊欄位)...");
                        
                        int r = startRowIdx;

                        foreach (string line in lines) {
                            // 如果貼上的列數超過現有的列數，自動產生新列
                            if (r >= workingDt.Rows.Count) {
                                workingDt.Rows.Add(workingDt.NewRow());
                            }

                            string[] cells = line.Split('\t');
                            int currentVisibleColIdx = visibleStartColIndex;

                            // 依照順序逐一填入「可見」的欄位
                            for (int i = 0; i < cells.Length; i++) {
                                // 如果剪貼簿的欄位數已經超過畫面上剩餘的可見欄位，就直接截斷不貼
                                if (currentVisibleColIdx >= visibleCols.Count) break; 

                                DataGridViewColumn dgvCol = visibleCols[currentVisibleColIdx];
                                string colName = dgvCol.Name;

                                // 🟢 嚴格防呆：只允許填入非唯讀、且不是 Id 和 附件檔案 的欄位
                                // (Id 被擋下，意味著如果是新列，Id就會維持空白，存檔時會自動產生新的，解決重複覆蓋問題)
                                if (!dgvCol.ReadOnly && colName != "Id" && colName != "附件檔案" && workingDt.Columns.Contains(colName)) {
                                    workingDt.Rows[r][colName] = cells[i].Trim().Trim('"');
                                }

                                currentVisibleColIdx++;
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

                // 更新畫面
                _dgv.Visible = false;
                UnfreezeAllColumns();
                _isApplyingWidths = true;
                _dgv.SuspendLayout();

                _dgv.DataSource = workingDt; 
                ApplyGridStyles(); 
                RestoreColumnOrder();
                ApplyFreezeState();

                _dgv.ResumeLayout(true);
                _dgv.Visible = true;
                _isApplyingWidths = false;
                _isBulkUpdating = false; 
            }
        }

        private void Dgv_KeyPress(object sender, KeyPressEventArgs e) 
        {
            if (_dgv.CurrentCell != null && !_dgv.CurrentCell.ReadOnly && !_dgv.IsCurrentCellInEditMode) {
                if (char.IsLetterOrDigit(e.KeyChar) || char.IsPunctuation(e.KeyChar) || char.IsSymbol(e.KeyChar) || char.IsWhiteSpace(e.KeyChar)) {
                    
                    string key = $"{_tableName}|{_dgv.Columns[_dgv.CurrentCell.ColumnIndex].Name}";
                    if (App_DropdownManager.DropdownCache.Keys.Any(k => k.StartsWith(key + "|")) || 
                        App_DropdownManager.MultiSelectCache.ContainsKey(key)) {
                        e.Handled = true;
                        return;
                    }

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
            if (e.Control is TextBox) { 
                TextBox txt = (TextBox)e.Control;
                txt.Multiline = true; 
                txt.KeyDown -= new KeyEventHandler(TextBox_KeyDown); 
                txt.KeyDown += new KeyEventHandler(TextBox_KeyDown); 
            }
        }

        private void Dgv_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            string colName = _dgv.Columns[e.ColumnIndex].Name;

            bool isSingleDropdown = App_DropdownManager.DropdownCache.Keys.Any(k => k.StartsWith($"{_tableName}|{colName}|"));
            bool isMultiDropdown = App_DropdownManager.MultiSelectCache.ContainsKey($"{_tableName}|{colName}");
            
            if ((isSingleDropdown || isMultiDropdown) && e.Value != null)
            {
                string valStr = e.Value.ToString().Trim();
                if (string.IsNullOrEmpty(valStr)) return;

                Rectangle rect = e.CellBounds;
                int imgSize = 24; 
                
                bool hasAnyIcon = false;
                List<Image> iconsToDraw = new List<Image>();

                if (isSingleDropdown)
                {
                    string prefix = $"{_tableName}|{colName}|";
                    foreach (var kvp in App_DropdownManager.DropdownCache) {
                        if (kvp.Key.StartsWith(prefix)) {
                            var match = kvp.Value.FirstOrDefault(d => d.Text == valStr);
                            if (match != null && !string.IsNullOrEmpty(match.IconBase64)) {
                                Image img = match.GetImage();
                                if (img != null) {
                                    iconsToDraw.Add(img);
                                    hasAnyIcon = true;
                                }
                                break;
                            }
                        }
                    }
                }
                else if (isMultiDropdown)
                {
                    string key = $"{_tableName}|{colName}";
                    var opts = valStr.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
                    
                    if (App_DropdownManager.MultiSelectCache.ContainsKey(key)) {
                        var defs = App_DropdownManager.MultiSelectCache[key];
                        foreach(var opt in opts) {
                            var match = defs.FirstOrDefault(d => d.Text == opt);
                            if (match != null && !string.IsNullOrEmpty(match.IconBase64)) {
                                Image img = match.GetImage();
                                if (img != null) {
                                    iconsToDraw.Add(img);
                                    hasAnyIcon = true;
                                }
                            }
                        }
                    }
                }

                if (hasAnyIcon)
                {
                    e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border | DataGridViewPaintParts.Focus | DataGridViewPaintParts.SelectionBackground);

                    int startX = rect.X + 6;
                    int imgY = rect.Y + (rect.Height - imgSize) / 2;

                    foreach(var img in iconsToDraw) {
                        e.Graphics.DrawImage(img, startX, imgY, imgSize, imgSize);
                        startX += imgSize + 4; 
                    }
                    
                    e.Handled = true; 
                }
            }
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

        private void TriggerMultiSelectDialog(int rowIndex, int colIndex)
        {
            if (rowIndex < 0 || colIndex < 0 || rowIndex >= _dgv.Rows.Count) return;
            
            string colName = _dgv.Columns[colIndex].Name;
            string multiKey = $"{_tableName}|{colName}";

            if (App_DropdownManager.MultiSelectCache.ContainsKey(multiKey))
            {
                var options = App_DropdownManager.MultiSelectCache[multiKey];
                string currentVal = _dgv[colIndex, rowIndex].Value?.ToString() ?? "";

                using (Form fMulti = new Form())
                {
                    fMulti.Text = $"複選組合：{colName}";
                    fMulti.Size = new Size(380, 500); 
                    fMulti.StartPosition = FormStartPosition.CenterParent;
                    fMulti.FormBorderStyle = FormBorderStyle.FixedDialog;
                    fMulti.MaximizeBox = false;
                    fMulti.MinimizeBox = false;
                    fMulti.BackColor = Color.White;

                    DataGridView dgvMulti = new DataGridView {
                        Dock = DockStyle.Fill,
                        BackgroundColor = Color.White,
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        AllowUserToResizeColumns = false,
                        AllowUserToResizeRows = false,
                        RowHeadersVisible = false,
                        ColumnHeadersVisible = false,
                        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                        CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                        Font = new Font("Microsoft JhengHei UI", 12F)
                    };
                    dgvMulti.RowTemplate.Height = 40; 

                    DataGridViewCheckBoxColumn chkCol = new DataGridViewCheckBoxColumn { Name = "Check", Width = 40, AutoSizeMode = DataGridViewAutoSizeColumnMode.None };
                    DataGridViewImageColumn imgCol = new DataGridViewImageColumn { Name = "Icon", Width = 40, AutoSizeMode = DataGridViewAutoSizeColumnMode.None, ImageLayout = DataGridViewImageCellLayout.Zoom };
                    DataGridViewTextBoxColumn txtCol = new DataGridViewTextBoxColumn { Name = "Text", ReadOnly = true };

                    dgvMulti.Columns.Add(chkCol);
                    dgvMulti.Columns.Add(imgCol);
                    dgvMulti.Columns.Add(txtCol);

                    string[] currentSelected = currentVal.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
                    
                    foreach (var opt in options)
                    {
                        int rIdx = dgvMulti.Rows.Add();
                        bool isChecked = currentSelected.Contains(opt.Text);
                        dgvMulti.Rows[rIdx].Cells["Check"].Value = isChecked;
                        dgvMulti.Rows[rIdx].Cells["Icon"].Value = opt.GetImage() ?? new Bitmap(1, 1); 
                        dgvMulti.Rows[rIdx].Cells["Text"].Value = opt.Text;
                        dgvMulti.Rows[rIdx].Tag = opt.Text; 
                    }

                    dgvMulti.CellClick += (s, ev) => {
                        if (ev.RowIndex >= 0) {
                            bool current = Convert.ToBoolean(dgvMulti.Rows[ev.RowIndex].Cells["Check"].Value);
                            dgvMulti.Rows[ev.RowIndex].Cells["Check"].Value = !current;
                        }
                    };

                    Button btnOk = new Button {
                        Text = "✔️ 確認並填入",
                        Dock = DockStyle.Bottom,
                        Height = 45,
                        BackColor = Color.SteelBlue,
                        ForeColor = Color.White,
                        Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                        Cursor = Cursors.Hand
                    };
                    
                    btnOk.Click += (s2, ev2) => {
                        List<string> selectedItems = new List<string>();
                        foreach (DataGridViewRow r in dgvMulti.Rows) {
                            if (Convert.ToBoolean(r.Cells["Check"].Value)) {
                                selectedItems.Add(r.Tag.ToString());
                            }
                        }
                        
                        _dgv[colIndex, rowIndex].Value = string.Join(Environment.NewLine, selectedItems);
                        _dgv.EndEdit();
                        fMulti.DialogResult = DialogResult.OK;
                    };

                    fMulti.Controls.Add(dgvMulti);
                    fMulti.Controls.Add(btnOk);
                    fMulti.ShowDialog(Form.ActiveForm);
                }
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e) 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < _dgv.Rows.Count && !_dgv.Rows[e.RowIndex].IsNewRow) 
            {
                string colName = _dgv.Columns[e.ColumnIndex].Name;

                if (colName.Contains("附件檔案")) {
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
                } 
                else if (App_DropdownManager.MultiSelectCache.ContainsKey($"{_tableName}|{colName}")) 
                {
                    TriggerMultiSelectDialog(e.RowIndex, e.ColumnIndex);
                } 
                else 
                {
                    string prefix = $"{_tableName}|{colName}|";
                    bool isSingleDropdown = App_DropdownManager.DropdownCache.Keys.Any(k => k.StartsWith(prefix));
                    bool isReferenceDropdown = App_DropdownManager.ReferenceCache.ContainsKey($"{_dbName}|{_tableName}|{colName}"); // 🟢 新增跨表判斷

                    if (isSingleDropdown || isReferenceDropdown)
                    {
                        ShowCustomDropdown(e.RowIndex, e.ColumnIndex);
                    }
                    else
                    {
                        _logic.OnCellClick(_dgv, e); 
                    }
                }
            }
        }

        private void Dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.RowIndex < _dgv.Rows.Count) 
            {
                string colName = _dgv.Columns[e.ColumnIndex].Name;
                if (App_DropdownManager.MultiSelectCache.ContainsKey($"{_tableName}|{colName}")) 
                {
                    TriggerMultiSelectDialog(e.RowIndex, e.ColumnIndex);
                }
            }
        }

        private void ShowCustomDropdown(int rowIndex, int colIndex)
        {
            string colName = _dgv.Columns[colIndex].Name;
            string parentColName = "";
            
            foreach (var kvp in App_DropdownManager.DropdownCache) {
                string[] parts = kvp.Key.Split('|');
                if (parts.Length == 4 && parts[0] == _tableName && parts[1] == colName && !string.IsNullOrEmpty(parts[2])) {
                    parentColName = parts[2]; break;
                }
            }

            List<DropdownItemDef> items = new List<DropdownItemDef>();

            // 🟢 新增：優先檢查是否為跨表參照
            string refKey = $"{_dbName}|{_tableName}|{colName}";
            if (App_DropdownManager.ReferenceCache.ContainsKey(refKey)) 
            {
                var refOpts = App_DropdownManager.GetReferenceOptions(_dbName, _tableName, colName);
                if (refOpts != null) {
                    foreach(var opt in refOpts) {
                        items.Add(new DropdownItemDef { Text = opt, IconBase64 = "" }); // 跨表暫無圖示
                    }
                }
            }
            else 
            {
                if (!string.IsNullOrEmpty(parentColName) && _dgv.Columns.Contains(parentColName)) {
                    string parentVal = _dgv.Rows[rowIndex].Cells[parentColName].Value?.ToString() ?? "";
                    string key = $"{_tableName}|{colName}|{parentColName}|{parentVal}";
                    if (App_DropdownManager.DropdownCache.ContainsKey(key)) {
                        items = App_DropdownManager.DropdownCache[key];
                    }
                } else {
                    string prefix = $"{_tableName}|{colName}|";
                    foreach(var kvp in App_DropdownManager.DropdownCache) {
                        if (kvp.Key.StartsWith(prefix)) {
                            items.AddRange(kvp.Value);
                        }
                    }
                }
            }

            var uniqueItems = items.GroupBy(x => x.Text).Select(g => g.First()).ToList();
            if (uniqueItems.Count == 0) return;

            ToolStripDropDown dropDown = new ToolStripDropDown();
            dropDown.Margin = Padding.Empty;
            dropDown.Padding = Padding.Empty;

            ListBox listBox = new ListBox {
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 36, 
                Font = new Font("Microsoft JhengHei UI", 12F),
                BorderStyle = BorderStyle.FixedSingle
            };

            Rectangle cellRect = _dgv.GetCellDisplayRectangle(colIndex, rowIndex, false);
            
            int maxTextWidth = 180;
            using (Graphics g = _dgv.CreateGraphics()) {
                foreach (var item in uniqueItems) {
                    int w = (int)g.MeasureString(item.Text, listBox.Font).Width;
                    if (w + 60 > maxTextWidth) maxTextWidth = w + 60; 
                }
            }

            int finalWidth = Math.Max(cellRect.Width, maxTextWidth);
            int finalHeight = uniqueItems.Count * 36 + 5;
            if (finalHeight > 300) finalHeight = 300; 

            listBox.Width = finalWidth;
            listBox.Height = finalHeight;

            listBox.DrawItem += (s, e) => {
                if (e.Index < 0) return;
                e.DrawBackground();

                var item = uniqueItems[e.Index];
                Brush textBrush = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ? Brushes.White : Brushes.Black;
                
                int imgSize = 24; 
                int textOffset = 8; 

                Image img = item.GetImage();
                if (img != null) {
                    int imgY = e.Bounds.Y + (e.Bounds.Height - imgSize) / 2;
                    e.Graphics.DrawImage(img, e.Bounds.X + 8, imgY, imgSize, imgSize);
                    textOffset = 40; 
                }

                e.Graphics.DrawString(item.Text, listBox.Font, textBrush, new PointF(e.Bounds.X + textOffset, e.Bounds.Y + 8));
                e.DrawFocusRectangle();
            };

            foreach (var item in uniqueItems) listBox.Items.Add(item);

            listBox.MouseClick += (s, e) => {
                if (listBox.SelectedIndex >= 0) {
                    _dgv[colIndex, rowIndex].Value = uniqueItems[listBox.SelectedIndex].Text;
                    _dgv.EndEdit();
                    dropDown.Close();
                }
            };

            listBox.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter && listBox.SelectedIndex >= 0) {
                    _dgv[colIndex, rowIndex].Value = uniqueItems[listBox.SelectedIndex].Text;
                    _dgv.EndEdit();
                    dropDown.Close();
                } else if (e.KeyCode == Keys.Escape) {
                    dropDown.Close();
                }
            };

            ToolStripControlHost host = new ToolStripControlHost(listBox);
            host.Margin = Padding.Empty;
            host.Padding = Padding.Empty;
            
            host.AutoSize = false;
            host.Size = new Size(finalWidth, finalHeight);
            
            dropDown.Items.Add(host);

            Point displayLocation = _dgv.PointToScreen(new Point(cellRect.Left, cellRect.Bottom));
            dropDown.Show(displayLocation);
            
            listBox.Focus();
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

            _isCascading = true;
            try {
                foreach (var kvp in App_DropdownManager.DropdownCache) {
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
