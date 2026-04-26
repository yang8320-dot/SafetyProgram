/// FILE: Safety_System/DataManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    public static class DataManager
    {
        private static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory;
        // 🟢 系統核心設定資料庫，取代所有 TXT 檔
        private static readonly string SysConfigDbPath = Path.Combine(AppDir, "SystemConfig.sqlite");

        public static string BasePath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");

        // 靜態建構子：確保系統啟動時配置資料庫存在
        static DataManager()
        {
            InitSysConfigDB();
        }

        private static void InitSysConfigDB()
        {
            string connStr = $"Data Source={SysConfigDbPath};Version=3;";
            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(conn))
                {
                    // 系統基礎變數 (取代 sys_config.txt, backup_config.txt)
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS SysSettings (KeyName TEXT PRIMARY KEY, KeyValue TEXT);";
                    cmd.ExecuteNonQuery();

                    // 防呆金鑰設定 (取代 table_keys.txt)
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS TableKeys (DbName TEXT, TableName TEXT, Col1 TEXT, Col2 TEXT, Col3 TEXT, Col4 TEXT, PRIMARY KEY(DbName, TableName));";
                    cmd.ExecuteNonQuery();

                    // 表格UI設定：顯示、寬度、順序 (取代 ColVisibility_...txt, ColWidths_...txt, ColOrder_...txt)
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS GridConfigs (DbName TEXT, TableName TEXT, ConfigType TEXT, ColName TEXT, ColValue TEXT, PRIMARY KEY(DbName, TableName, ConfigType, ColName));";
                    cmd.ExecuteNonQuery();

                    // 水資源自訂公式 (取代 WaterCustomStats.txt)
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS CustomStats (Module TEXT, StatName TEXT, Unit TEXT, Formula TEXT, PRIMARY KEY(Module, StatName));";
                    cmd.ExecuteNonQuery();

                    // 資料同步規則設定表
                    cmd.CommandText = @"CREATE TABLE IF NOT EXISTS SyncRules (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        SrcDb TEXT, SrcTable TEXT, SrcMatchCol TEXT, SrcSyncCol TEXT,
                                        TgtDb TEXT, TgtTable TEXT, TgtMatchCol TEXT, TgtSyncCol TEXT);";
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // ==============================================================
        // 🟢 系統設定讀寫區 (SysSettings)
        // ==============================================================
        public static string GetSysSetting(string key, string defaultValue = "")
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT KeyValue FROM SysSettings WHERE KeyName=@K", conn)) {
                        cmd.Parameters.AddWithValue("@K", key);
                        var res = cmd.ExecuteScalar();
                        return res != null ? res.ToString() : defaultValue;
                    }
                }
            } catch { return defaultValue; }
        }

        public static void SetSysSetting(string key, string value)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("INSERT INTO SysSettings (KeyName, KeyValue) VALUES (@K, @V) ON CONFLICT(KeyName) DO UPDATE SET KeyValue=@V", conn)) {
                        cmd.Parameters.AddWithValue("@K", key);
                        cmd.Parameters.AddWithValue("@V", value);
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        public static void LoadConfig()
        {
            string savedPath = GetSysSetting("BasePath", "");
            if (!string.IsNullOrEmpty(savedPath)) BasePath = savedPath;
            if (!Directory.Exists(BasePath)) Directory.CreateDirectory(BasePath);
        }

        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            SetSysSetting("BasePath", newPath);
            if (!Directory.Exists(BasePath)) Directory.CreateDirectory(BasePath);
        }

        // ==============================================================
        // 🟢 DataGridView UI 設定記憶區 (GridConfigs)
        // ==============================================================
        public static void SaveGridConfig(string dbName, string tableName, string configType, string colName, string value)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = "INSERT INTO GridConfigs (DbName, TableName, ConfigType, ColName, ColValue) VALUES (@DB, @TB, @Type, @Col, @Val) ON CONFLICT(DbName, TableName, ConfigType, ColName) DO UPDATE SET ColValue=@Val";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                        cmd.Parameters.AddWithValue("@Type", configType); cmd.Parameters.AddWithValue("@Col", colName); cmd.Parameters.AddWithValue("@Val", value);
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        public static Dictionary<string, string> LoadGridConfig(string dbName, string tableName, string configType)
        {
            var dict = new Dictionary<string, string>();
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT ColName, ColValue FROM GridConfigs WHERE DbName=@DB AND TableName=@TB AND ConfigType=@Type", conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName); cmd.Parameters.AddWithValue("@Type", configType);
                        using (var reader = cmd.ExecuteReader()) {
                            while (reader.Read()) { dict[reader["ColName"].ToString()] = reader["ColValue"].ToString(); }
                        }
                    }
                }
            } catch { }
            return dict;
        }

        public static void ClearGridConfig(string dbName, string tableName, string configType)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("DELETE FROM GridConfigs WHERE DbName=@DB AND TableName=@TB AND ConfigType=@Type", conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName); cmd.Parameters.AddWithValue("@Type", configType);
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        // ==============================================================
        // 🟢 資料表資料庫連線與基本操作
        // ==============================================================
        private static string GetConnString(string dbName)
        {
            string fullPath = Path.Combine(BasePath, dbName + ".sqlite");
            return string.Format("Data Source={0};Version=3;Default Timeout=30;Pooling=True;Max Pool Size=100;", fullPath);
        }

        private static void ExecuteWithRetry(string dbName, Action<SQLiteConnection> dbAction)
        {
            int maxRetries = 5;
            for (int i = 1; i <= maxRetries; i++) {
                try {
                    using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                        conn.Open(); 
                        dbAction(conn); 
                        return;
                    }
                } catch (SQLiteException ex) when (ex.ResultCode == SQLiteErrorCode.Locked || ex.ResultCode == SQLiteErrorCode.Busy) {
                    if (i == maxRetries) throw;
                    Thread.Sleep(100 * i);
                }
            }
        }

        // ==============================================================
        // 🟢 防重寫欄位操作 (TableKeys)
        // ==============================================================
        public static (string col1, string col2, string col3, string col4) GetTableKeys(string dbName, string tableName)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT Col1, Col2, Col3, Col4 FROM TableKeys WHERE DbName=@DB AND TableName=@TB", conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                        using (var reader = cmd.ExecuteReader()) {
                            if (reader.Read()) {
                                return (reader["Col1"].ToString(), reader["Col2"].ToString(), reader["Col3"].ToString(), reader["Col4"].ToString());
                            }
                        }
                    }
                }
            } catch { }
            return ("", "", "", "");
        }

        public static void SaveTableKeys(string dbName, string tableName, string col1, string col2, string col3, string col4)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    string sql = "INSERT INTO TableKeys (DbName, TableName, Col1, Col2, Col3, Col4) VALUES (@DB, @TB, @C1, @C2, @C3, @C4) ON CONFLICT(DbName, TableName) DO UPDATE SET Col1=@C1, Col2=@C2, Col3=@C3, Col4=@C4";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); cmd.Parameters.AddWithValue("@TB", tableName);
                        cmd.Parameters.AddWithValue("@C1", col1); cmd.Parameters.AddWithValue("@C2", col2);
                        cmd.Parameters.AddWithValue("@C3", col3); cmd.Parameters.AddWithValue("@C4", col4);
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        // ==============================================================
        // 🟢 智慧型批次存檔與同步機制 (BulkSaveTable)
        // ==============================================================
        public static bool BulkSaveTable(string dbName, string tableName, DataTable dt)
        {
            try {
                using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                    conn.Open();
                    using (var trans = conn.BeginTransaction()) {
                        var keys = GetTableKeys(dbName, tableName);

                        List<string> activeKeys = new List<string>();
                        if (!string.IsNullOrEmpty(keys.col1) && dt.Columns.Contains(keys.col1)) activeKeys.Add(keys.col1);
                        if (!string.IsNullOrEmpty(keys.col2) && dt.Columns.Contains(keys.col2)) activeKeys.Add(keys.col2);
                        if (!string.IsNullOrEmpty(keys.col3) && dt.Columns.Contains(keys.col3)) activeKeys.Add(keys.col3);
                        if (!string.IsNullOrEmpty(keys.col4) && dt.Columns.Contains(keys.col4)) activeKeys.Add(keys.col4);

                        foreach (DataRow row in dt.Rows) {
                            if (row.RowState == DataRowState.Deleted) continue;

                            foreach (DataColumn col in dt.Columns) {
                                if (col.ColumnName.Contains("日期") && row[col] != DBNull.Value) {
                                    if (DateTime.TryParse(row[col].ToString(), out DateTime d))
                                        row[col] = d.ToString("yyyy-MM-dd");
                                }
                            }

                            int existingId = -1;
                            
                            if (activeKeys.Count > 0) {
                                List<string> whereClauses = new List<string>();
                                foreach (var k in activeKeys) {
                                    string safeKey = k.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                    whereClauses.Add($"IFNULL([{k}], '') = IFNULL(@{safeKey}, '')");
                                }

                                string qCheck = $"SELECT Id FROM [{tableName}] WHERE " + string.Join(" AND ", whereClauses);

                                using (var cmdCheck = new SQLiteCommand(qCheck, conn, trans)) {
                                    foreach (var k in activeKeys) {
                                        string safeKey = k.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                        object val = row[k] != DBNull.Value ? (object)row[k].ToString().Trim() : DBNull.Value;
                                        cmdCheck.Parameters.AddWithValue("@" + safeKey, val);
                                    }
                                    var res = cmdCheck.ExecuteScalar();
                                    if (res != null && res != DBNull.Value) existingId = Convert.ToInt32(res);
                                }
                            }

                            bool isUpdate = (existingId != -1) || (dt.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0);
                            using (var cmd = new SQLiteCommand(conn)) {
                                cmd.Transaction = trans;
                                List<string> sqlParts = new List<string>();
                                
                                if (isUpdate) {
                                    int targetId = existingId != -1 ? existingId : Convert.ToInt32(row["Id"]);
                                    string sql = $"UPDATE [{tableName}] SET ";
                                    foreach (DataColumn col in dt.Columns) {
                                        if (col.ColumnName == "Id") continue;
                                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                        sqlParts.Add($"[{col.ColumnName}]=@{safeParamName}");
                                        object val = row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value;
                                        cmd.Parameters.AddWithValue("@" + safeParamName, val);
                                    }
                                    cmd.CommandText = sql + string.Join(", ", sqlParts) + " WHERE Id=" + targetId;
                                } else {
                                    List<string> colNames = new List<string>();
                                    List<string> paramNames = new List<string>();
                                    foreach (DataColumn col in dt.Columns) {
                                        if (col.ColumnName == "Id") continue;
                                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                                        colNames.Add($"[{col.ColumnName}]");
                                        paramNames.Add($"@{safeParamName}");
                                        object val = row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value;
                                        cmd.Parameters.AddWithValue("@" + safeParamName, val);
                                    }
                                    cmd.CommandText = $"INSERT INTO [{tableName}] ({string.Join(", ", colNames)}) VALUES ({string.Join(", ", paramNames)})";
                                }
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                
                // 🟢 儲存成功後，呼叫全域同步引擎
                RunSyncEngine(dbName, tableName);

                return true;
            } catch (Exception ex) {
                MessageBox.Show("批次儲存失敗：" + ex.Message);
                return false;
            }
        }

        public static void UpsertRecord(string dbName, string tableName, DataRow row)
        {
            ExecuteWithRetry(dbName, conn => {
                bool isUpdate = row.Table.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0;
                var cmd = new SQLiteCommand(conn);
                if (isUpdate) {
                    List<string> sets = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                        sets.Add($"[{col.ColumnName}]=@{safeParamName}");
                        cmd.Parameters.AddWithValue("@" + safeParamName, row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value);
                    }
                    cmd.CommandText = $"UPDATE [{tableName}] SET {string.Join(", ", sets)} WHERE Id=" + row["Id"];
                } else {
                    List<string> c = new List<string>();
                    List<string> v = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id") continue;
                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "");
                        c.Add($"[{col.ColumnName}]"); 
                        v.Add($"@{safeParamName}");
                        cmd.Parameters.AddWithValue("@" + safeParamName, row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value);
                    }
                    cmd.CommandText = $"INSERT INTO [{tableName}] ({string.Join(", ", c)}) VALUES ({string.Join(", ", v)})";
                }
                cmd.ExecuteNonQuery();
            });

            // 🟢 儲存成功後，呼叫全域同步引擎
            RunSyncEngine(dbName, tableName);
        }

        public static void DeleteRecord(string dbName, string tableName, int id) 
        {
            ExecuteWithRetry(dbName, conn => {
                using (var cmd = new SQLiteCommand($"DELETE FROM [{tableName}] WHERE Id=@Id", conn)) { cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery(); }
            });

            // 🟢 刪除成功後，呼叫全域同步引擎重新計算
            RunSyncEngine(dbName, tableName);
        }

        // ==============================================================
        // 🟢 全域自動同步引擎 (SyncEngine)
        // ==============================================================
        public static void RunSyncEngine(string srcDb, string srcTable)
        {
            try
            {
                DataTable rules = new DataTable();
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;"))
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT * FROM SyncRules WHERE SrcDb=@DB AND SrcTable=@TB", conn))
                    {
                        cmd.Parameters.AddWithValue("@DB", srcDb);
                        cmd.Parameters.AddWithValue("@TB", srcTable);
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(rules);
                    }
                }

                if (rules.Rows.Count == 0) return;

                foreach (DataRow rule in rules.Rows)
                {
                    string srcMatchCol = rule["SrcMatchCol"].ToString();
                    string srcSyncCol = rule["SrcSyncCol"].ToString();
                    string tgtDb = rule["TgtDb"].ToString();
                    string tgtTable = rule["TgtTable"].ToString();
                    string tgtMatchCol = rule["TgtMatchCol"].ToString();
                    string tgtSyncCol = rule["TgtSyncCol"].ToString();

                    // 取得來源的所有聚合統計資料
                    DataTable aggregatedData = new DataTable();
                    ExecuteWithRetry(srcDb, connSrc => {
                        // 判斷是否為「日」同步到「月」(假設日格式為 YYYY-MM-DD，月為 YYYY-MM)
                        bool isDateToMonth = srcMatchCol.Contains("日") && tgtMatchCol.Contains("月");
                        
                        string groupColSql = isDateToMonth ? $"substr([{srcMatchCol}], 1, 7)" : $"[{srcMatchCol}]";
                        
                        // 將來源字串數值(可能含逗號)轉為浮點數進行加總
                        string sql = $@"
                            SELECT {groupColSql} as MatchVal, SUM(CAST(REPLACE([{srcSyncCol}], ',', '') AS REAL)) as SyncVal 
                            FROM [{srcTable}] 
                            WHERE [{srcMatchCol}] IS NOT NULL AND [{srcMatchCol}] != '' 
                            GROUP BY {groupColSql}";

                        using (var cmd = new SQLiteCommand(sql, connSrc))
                        {
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(aggregatedData);
                        }
                    });

                    // 確保目標資料表存在該欄位 (如果沒有，嘗試建立)
                    var targetCols = GetColumnNames(tgtDb, tgtTable);
                    if (targetCols.Count > 0 && !targetCols.Contains(tgtSyncCol))
                    {
                        AddColumn(tgtDb, tgtTable, tgtSyncCol);
                    }

                    // 執行目標資料庫的 Upsert 寫入
                    ExecuteWithRetry(tgtDb, connTgt => {
                        using (var trans = connTgt.BeginTransaction())
                        {
                            foreach (DataRow row in aggregatedData.Rows)
                            {
                                string matchVal = row["MatchVal"].ToString();
                                string syncVal = row["SyncVal"].ToString();

                                // 確認目標表是否存在該筆 MatchVal
                                string checkSql = $"SELECT Id FROM [{tgtTable}] WHERE [{tgtMatchCol}] = @M";
                                int tgtId = -1;
                                using (var checkCmd = new SQLiteCommand(checkSql, connTgt, trans))
                                {
                                    checkCmd.Parameters.AddWithValue("@M", matchVal);
                                    var res = checkCmd.ExecuteScalar();
                                    if (res != null && res != DBNull.Value) tgtId = Convert.ToInt32(res);
                                }

                                using (var cmd = new SQLiteCommand(connTgt))
                                {
                                    cmd.Transaction = trans;
                                    if (tgtId != -1)
                                    {
                                        cmd.CommandText = $"UPDATE [{tgtTable}] SET [{tgtSyncCol}] = @V WHERE Id = @Id";
                                        cmd.Parameters.AddWithValue("@V", syncVal);
                                        cmd.Parameters.AddWithValue("@Id", tgtId);
                                    }
                                    else
                                    {
                                        // 若無資料，則新增該月份的紀錄
                                        cmd.CommandText = $"INSERT INTO [{tgtTable}] ([{tgtMatchCol}], [{tgtSyncCol}]) VALUES (@M, @V)";
                                        cmd.Parameters.AddWithValue("@M", matchVal);
                                        cmd.Parameters.AddWithValue("@V", syncVal);
                                    }
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            trans.Commit();
                        }
                    });
                }
            }
            catch { }
        }

        // ==============================================================
        // 🟢 基礎架構輔助操作
        // ==============================================================
        public static List<string> GetColumnNames(string dbName, string tableName)
        {
            var cols = new List<string>();
            try {
                ExecuteWithRetry(dbName, conn => {
                    using (var cmd = new SQLiteCommand($"PRAGMA table_info([{tableName}])", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) cols.Add(r["name"].ToString()); }
                });
            } catch { } 
            return cols;
        }

        public static void InitTable(string dbName, string tableName, string createSql) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand(createSql, conn)) cmd.ExecuteNonQuery();
        });

        public static DataTable GetTableData(string dbName, string tableName, string dateCol, string start, string end)
        {
            DataTable dt = new DataTable();
            ExecuteWithRetry(dbName, conn => {
                string q = string.IsNullOrEmpty(start) ? $"SELECT * FROM [{tableName}]" : $"SELECT * FROM [{tableName}] WHERE [{dateCol}] BETWEEN @s AND @e";
                using (var cmd = new SQLiteCommand(q, conn)) {
                    if (!string.IsNullOrEmpty(start)) { cmd.Parameters.AddWithValue("@s", start); cmd.Parameters.AddWithValue("@e", end); }
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            });
            return dt;
        }

        public static DataTable GetLatestRecords(string dbName, string tableName, int limit = 30)
        {
            DataTable dt = new DataTable();
            ExecuteWithRetry(dbName, conn => {
                string q = limit > 0 ? $"SELECT * FROM (SELECT * FROM [{tableName}] ORDER BY Id DESC LIMIT @limit) sub ORDER BY Id ASC" 
                                     : $"SELECT * FROM [{tableName}] ORDER BY Id ASC";
                using (var cmd = new SQLiteCommand(q, conn)) {
                    if (limit > 0) cmd.Parameters.AddWithValue("@limit", limit);
                    using (var da = new SQLiteDataAdapter(cmd)) da.Fill(dt);
                }
            });
            return dt;
        }

        public static void AddColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] ADD COLUMN [{colName}] TEXT", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void RenameColumn(string dbName, string tableName, string oldN, string newN) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] RENAME COLUMN [{oldN}] TO [{newN}]", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] DROP COLUMN [{colName}]", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropTable(string dbName, string tableName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"DROP TABLE IF EXISTS [{tableName}]", conn)) { cmd.ExecuteNonQuery(); }
        });
    }
}
