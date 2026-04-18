/*
 * 檔案功能：SQLite 資料庫全域管理器 (負責連線與初始化資料表)
 * 對應選單名稱：全域共用
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表名稱：TodoList, Shortcuts
 */

using System;
using System.Data.SQLite;
using System.IO;

public static class DatabaseManager
{
    private static string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MainDB.sqlite");
    private static string connectionString = $"Data Source={dbPath};Version=3;";

    public static void InitializeDatabase()
    {
        if (!File.Exists(dbPath))
        {
            SQLiteConnection.CreateFile(dbPath);
        }

        using (var conn = GetConnection())
        {
            conn.Open();
            
            // 1. 建立 TodoList 資料表 (待辦事項)
            string createTodoTable = @"
                CREATE TABLE IF NOT EXISTS TodoList (
                    Id TEXT PRIMARY KEY,
                    ListName TEXT NOT NULL,
                    Content TEXT NOT NULL,
                    CreatedDate TEXT NOT NULL
                )";
            using (var cmd = new SQLiteCommand(createTodoTable, conn)) { cmd.ExecuteNonQuery(); }

            // 2. 建立 Shortcuts 資料表 (常用捷徑) - 新增 SortOrder 負責記憶拖曳排序
            string createShortcutsTable = @"
                CREATE TABLE IF NOT EXISTS Shortcuts (
                    Id TEXT PRIMARY KEY,
                    Name TEXT NOT NULL,
                    TargetPath TEXT NOT NULL,
                    SortOrder INTEGER NOT NULL
                )";
            using (var cmd = new SQLiteCommand(createShortcutsTable, conn)) { cmd.ExecuteNonQuery(); }

            // (未來會在此處繼續加入 FileWatcher, RecurringTasks 等資料表)
        }
    }

    public static SQLiteConnection GetConnection()
    {
        return new SQLiteConnection(connectionString);
    }
}
