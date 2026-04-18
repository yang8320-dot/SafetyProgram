/*
 * 檔案功能：SQLite 資料庫全域管理器 (最終完整版)
 * 對應資料庫名稱：MainDB.sqlite
 * 資料表：TodoList, Shortcuts, FileWatcher, RecurringTasks, GlobalSettings
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
            
            // 1. 待辦事項 (TodoList)
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS TodoList (Id TEXT PRIMARY KEY, ListName TEXT NOT NULL, Content TEXT NOT NULL, CreatedDate TEXT NOT NULL)", conn)) { cmd.ExecuteNonQuery(); }

            // 2. 常用捷徑 (Shortcuts)
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS Shortcuts (Id TEXT PRIMARY KEY, Name TEXT NOT NULL, TargetPath TEXT NOT NULL, SortOrder INTEGER NOT NULL)", conn)) { cmd.ExecuteNonQuery(); }

            // 3. 檔案監控 (FileWatcher)
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS FileWatcher (SourcePath TEXT PRIMARY KEY, TargetPath TEXT NOT NULL, SyncMethod TEXT, Frequency TEXT, Depth TEXT, SyncMode TEXT, Retention TEXT, CustomName TEXT)", conn)) { cmd.ExecuteNonQuery(); }

            // 4. 週期任務 (RecurringTasks)
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS RecurringTasks (Id TEXT PRIMARY KEY, Name TEXT NOT NULL, MonthStr TEXT, DateStr TEXT, TimeStr TEXT, LastTriggeredDate TEXT, Note TEXT, TaskType TEXT)", conn)) { cmd.ExecuteNonQuery(); }

            // 5. 全域設定 (GlobalSettings - 供週期任務等全域變數使用)
            using (var cmd = new SQLiteCommand(@"CREATE TABLE IF NOT EXISTS GlobalSettings (SettingKey TEXT PRIMARY KEY, SettingValue TEXT)", conn)) { cmd.ExecuteNonQuery(); }
        }
    }

    public static SQLiteConnection GetConnection()
    {
        return new SQLiteConnection(connectionString);
    }
}
