// 在 UI_MainForm.cs 中的 SyncTasks 方法內，請確保如下修改：

private async Task SyncTasks()
{
    UpdateStatus("正在同步 Google Tasks...");
    try {
        // 明確指定回傳的是 Google API 的 Task 列表
        var items = await _service.GetAllTasksAsync(); 
        
        _taskView.Items.Clear();
        
        // 這裡使用 var 或明確寫出類型
        foreach (var t in items) {
            var lvi = new ListViewItem(t.Status == "completed" ? "[V]" : "[ ]");
            lvi.SubItems.Add(t.Title ?? "無標題");
            lvi.SubItems.Add(t.Id ?? "無 ID");
            _taskView.Items.Add(lvi);
        }
        UpdateStatus("同步完成。");
    } catch (Exception ex) { 
        MessageBox.Show("同步錯誤: " + ex.Message); 
        UpdateStatus("同步失敗"); 
    }
}
