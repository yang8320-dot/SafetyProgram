/*
 * 檔案功能：SQLite 資料庫全域管理器 (負責連線與初始化資料表)
 * 對應選單名稱：全域共用
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表名稱：TodoList (目前先初始化待辦事項，後續模組轉換時會加入其他表)
 */

using System;
using System.Data.SQLite;
using System.IO;

public static class DatabaseManager
{
    // 將資料庫檔案建立在執行檔同一個目錄下
    private static string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MainDB.sqlite");
    
    // SQLite 連線字串
    private static string connectionString = $"Data Source={dbPath};Version=3;";

    /// <summary>
    /// 初始化資料庫與所有必要的資料表
    /// </summary>
    public static void InitializeDatabase()
    {
        // 若檔案不存在，自動建立 SQLite 實體檔案
        if (!File.Exists(dbPath))
        {
            SQLiteConnection.CreateFile(dbPath);
        }

        using (var conn = GetConnection())
        {
            conn.Open();
            
            // 建立 TodoList 資料表
            // ListName 欄位用來區分是 "todo(待辦)" 還是 "plan(計畫)"
            string createTodoTable = @"
                CREATE TABLE IF NOT EXISTS TodoList (
                    Id TEXT PRIMARY KEY,
                    ListName TEXT NOT NULL,
                    Content TEXT NOT NULL,
                    CreatedDate TEXT NOT NULL
                )";

            using (var cmd = new SQLiteCommand(createTodoTable, conn))
            {
                cmd.ExecuteNonQuery();
            }

            // (未來轉換其他模組時，會在此處繼續加入 Shortcuts, FileWatcher 等資料表)
        }
    }

    /// <summary>
    /// 取得一個新的 SQLite 連線
    /// </summary>
    public static SQLiteConnection GetConnection()
    {
        return new SQLiteConnection(connectionString);
    }
}
