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
        private static readonly string SysConfigDbPath = Path.Combine(AppDir, "SystemConfig.sqlite");

        public static string BasePath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");

        [ThreadStatic]
        private static bool _isSyncing = false;

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
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS SysSettings (KeyName TEXT PRIMARY KEY, KeyValue TEXT);";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS TableKeys (DbName TEXT, TableName TEXT, Col1 TEXT, Col2 TEXT, Col3 TEXT, Col4 TEXT, PRIMARY KEY(DbName, TableName));";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS GridConfigs (DbName TEXT, TableName TEXT, ConfigType TEXT, ColName TEXT, ColValue TEXT, PRIMARY KEY(DbName, TableName, ConfigType, ColName));";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS CustomStats (Module TEXT, StatName TEXT, Unit TEXT, Formula TEXT, PRIMARY KEY(Module, StatName));";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = @"CREATE TABLE IF NOT EXISTS SyncRules (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        SrcDb TEXT, SrcTable TEXT, SrcMatchCol TEXT, SrcSyncCol TEXT,
                                        TgtDb TEXT, TgtTable TEXT, TgtMatchCol TEXT, TgtSyncCol TEXT,
                                        SyncType TEXT DEFAULT '單向同步');";
                    cmd.ExecuteNonQuery();

                    // 支援自訂選單擴充表
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS CustomMenus (Id INTEGER PRIMARY KEY AUTOINCREMENT, [分類] TEXT, [資料庫名] TEXT, [資料表名] TEXT);";
                    cmd.ExecuteNonQuery();
                }
            }
        }

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
            RunSyncEngine(dbName, tableName);
        }

        public static void DeleteRecord(string dbName, string tableName, int id) 
        {
            ExecuteWithRetry(dbName, conn => {
                using (var cmd = new SQLiteCommand($"DELETE FROM [{tableName}] WHERE Id=@Id", conn)) { cmd.Parameters.AddWithValue("@Id", id); cmd.ExecuteNonQuery(); }
            });
            RunSyncEngine(dbName, tableName);
        }

        // 全域自動同步引擎
        public static void RunSyncEngine(string triggerDb, string triggerTable)
        {
            if (_isSyncing) return;

            try
            {
                _isSyncing = true;
                DataTable rules = new DataTable();
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;"))
                {
                    conn.Open();
                    string sqlFetch = "SELECT * FROM SyncRules WHERE (SrcDb=@DB AND SrcTable=@TB) OR (TgtDb=@DB AND TgtTable=@TB AND SyncType='雙向同步')";
                    using (var cmd = new SQLiteCommand(sqlFetch, conn))
                    {
                        cmd.Parameters.AddWithValue("@DB", triggerDb);
                        cmd.Parameters.AddWithValue("@TB", triggerTable);
                        using (var da = new SQLiteDataAdapter(cmd)) da.Fill(rules);
                    }
                }

                if (rules.Rows.Count == 0) return;

                foreach (DataRow rule in rules.Rows)
                {
                    bool isReverse = (rule["TgtDb"].ToString() == triggerDb && rule["TgtTable"].ToString() == triggerTable);

                    string actualSrcDb = isReverse ? rule["TgtDb"].ToString() : rule["SrcDb"].ToString();
                    string actualSrcTable = isReverse ? rule["TgtTable"].ToString() : rule["SrcTable"].ToString();
                    string actualSrcMatchCol = isReverse ? rule["TgtMatchCol"].ToString() : rule["SrcMatchCol"].ToString();
                    string actualSrcSyncCol = isReverse ? rule["TgtSyncCol"].ToString() : rule["SrcSyncCol"].ToString();

                    string actualTgtDb = isReverse ? rule["SrcDb"].ToString() : rule["TgtDb"].ToString();
                    string actualTgtTable = isReverse ? rule["SrcTable"].ToString() : rule["TgtTable"].ToString();
                    string actualTgtMatchCol = isReverse ? rule["SrcMatchCol"].ToString() : rule["TgtMatchCol"].ToString();
                    string actualTgtSyncCol = isReverse ? rule["SrcSyncCol"].ToString() : rule["TgtSyncCol"].ToString();

                    DataTable aggregatedData = new DataTable();
                    ExecuteWithRetry(actualSrcDb, connSrc => {
                        bool isGrouping = (actualSrcMatchCol != actualTgtMatchCol) && (actualSrcMatchCol.Contains("日") && actualTgtMatchCol.Contains("月"));
                        
                        string sql = "";
                        if (isGrouping)
                        {
                            string groupColSql = $"substr([{actualSrcMatchCol}], 1, 7)";
                            sql = $@"
                                SELECT {groupColSql} as MatchVal, SUM(CAST(REPLACE(IFNULL([{actualSrcSyncCol}], '0'), ',', '') AS REAL)) as SyncVal 
                                FROM [{actualSrcTable}] 
                                WHERE [{actualSrcMatchCol}] IS NOT NULL AND [{actualSrcMatchCol}] != '' 
                                GROUP BY {groupColSql}";
                        }
                        else
                        {
                            sql = $@"
                                SELECT [{actualSrcMatchCol}] as MatchVal, [{actualSrcSyncCol}] as SyncVal 
                                FROM [{actualSrcTable}] 
                                WHERE [{actualSrcMatchCol}] IS NOT NULL AND [{actualSrcMatchCol}] != ''";
                        }

                        using (var cmd = new SQLiteCommand(sql, connSrc))
                        {
                            using (var da = new SQLiteDataAdapter(cmd)) da.Fill(aggregatedData);
                        }
                    });

                    var targetCols = GetColumnNames(actualTgtDb, actualTgtTable);
                    if (targetCols.Count > 0 && !targetCols.Contains(actualTgtSyncCol))
                    {
                        AddColumn(actualTgtDb, actualTgtTable, actualTgtSyncCol);
                    }

                    ExecuteWithRetry(actualTgtDb, connTgt => {
                        using (var trans = connTgt.BeginTransaction())
                        {
                            foreach (DataRow row in aggregatedData.Rows)
                            {
                                string matchVal = row["MatchVal"].ToString();
                                string syncVal = row["SyncVal"].ToString();

                                string checkSql = $"SELECT Id FROM [{actualTgtTable}] WHERE [{actualTgtMatchCol}] = @M";
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
                                        cmd.CommandText = $"UPDATE [{actualTgtTable}] SET [{actualTgtSyncCol}] = @V WHERE Id = @Id";
                                        cmd.Parameters.AddWithValue("@V", syncVal);
                                        cmd.Parameters.AddWithValue("@Id", tgtId);
                                    }
                                    else
                                    {
                                        cmd.CommandText = $"INSERT INTO [{actualTgtTable}] ([{actualTgtMatchCol}], [{actualTgtSyncCol}]) VALUES (@M, @V)";
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
            finally
            {
                _isSyncing = false; 
            }
        }

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

        // 🟢 支援重新命名資料表功能
        public static void RenameTable(string dbName, string oldName, string newName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{oldName}] RENAME TO [{newName}]", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropColumn(string dbName, string tableName, string colName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] DROP COLUMN [{colName}]", conn)) { cmd.ExecuteNonQuery(); }
        });

        public static void DropTable(string dbName, string tableName) => ExecuteWithRetry(dbName, conn => {
            using (var cmd = new SQLiteCommand($"DROP TABLE IF EXISTS [{tableName}]", conn)) { cmd.ExecuteNonQuery(); }
        });
    }
}
