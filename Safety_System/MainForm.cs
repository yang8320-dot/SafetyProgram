// 請將此段邏輯更新至 MainForm.cs 中的 BuildMenu 方法內

private void BuildMenu()
{
    // ... 前方的首頁、工安、環保等選單維持不變 ...

    // 5. 設定
    var menuSettings = new ToolStripMenuItem("設定");
    var itemDbConfig = new ToolStripMenuItem("資料庫設定");
    
    // 關鍵：點擊後以 Show() 開啟，這樣不會卡住主視窗
    itemDbConfig.Click += (s, e) => {
        App_DbConfig configForm = new App_DbConfig();
        configForm.Show(); 
    };

    menuSettings.DropDownItems.Add(itemDbConfig);

    _mainMenu.Items.AddRange(new ToolStripItem[] {
        menuHome, menuSafety, menuEnv, menuFire, 
        new ToolStripMenuItem("ESG"), new ToolStripMenuItem("溫盤"), 
        menuSettings
    });
}
