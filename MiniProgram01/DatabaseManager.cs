/*
 * 檔案功能：SQLite 資料庫全域管理器 (補回待辦事項 Color 與 Note 欄位)
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表：TodoList, Shortcuts, FileWatcher, RecurringTasks, GlobalSettings
 */

/*
 * 檔案功能：SQLite 資料庫全域管理器 (修正單一執行檔路徑問題)
 */

using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms; // 新增這行

public static class DatabaseManager
{
    // 修正：改用 Application.StartupPath 確保在單一執行檔模式下路徑正確
    private static string dbPath = Path.Combine(Application.StartupPath, "MainDB.sqlite");
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
            
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS TodoList (Id TEXT PRIMARY KEY, ListName TEXT NOT NULL, Content TEXT NOT NULL, CreatedDate TEXT NOT NULL, Color TEXT DEFAULT 'Black', Note TEXT DEFAULT '')", conn)) { cmd.ExecuteNonQuery(); }
            try { using (var cmd = new SQLiteCommand("ALTER TABLE TodoList ADD COLUMN Color TEXT DEFAULT 'Black'", conn)) { cmd.ExecuteNonQuery(); } } catch { }
            try { using (var cmd = new SQLiteCommand("ALTER TABLE TodoList ADD COLUMN Note TEXT DEFAULT ''", conn)) { cmd.ExecuteNonQuery(); } } catch { }
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS Shortcuts (Id TEXT PRIMARY KEY, Name TEXT NOT NULL, TargetPath TEXT NOT NULL, SortOrder INTEGER NOT NULL)", conn)) { cmd.ExecuteNonQuery(); }
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS FileWatcher (SourcePath TEXT PRIMARY KEY, TargetPath TEXT NOT NULL, SyncMethod TEXT, Frequency TEXT, Depth TEXT, SyncMode TEXT, Retention TEXT, CustomName TEXT)", conn)) { cmd.ExecuteNonQuery(); }
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS RecurringTasks (Id TEXT PRIMARY KEY, Name TEXT NOT NULL, MonthStr TEXT, DateStr TEXT, TimeStr TEXT, LastTriggeredDate TEXT, Note TEXT, TaskType TEXT)", conn)) { cmd.ExecuteNonQuery(); }
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS GlobalSettings (SettingKey TEXT PRIMARY KEY, SettingValue TEXT)", conn)) { cmd.ExecuteNonQuery(); }
        }
    }

    public static SQLiteConnection GetConnection()
    {
        return new SQLiteConnection(connectionString);
    }
}
