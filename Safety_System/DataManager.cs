/// FILE: Safety_System/DataManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Safety_System
{
    // 🟢 新增：支援多區間公式與「公式類型(數學/文字)」的資料模型
    public class ColumnFormulaDef {
        public int Id { get; set; }
        public string TargetCol { get; set; }
        public string MatchCol { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string FormulaType { get; set; } // 🟢 新增屬性：區分「數學運算」或「文字組合」
        public string Formula { get; set; }
    }

    public static class DataManager
    {
        private static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string SysConfigDbPath = Path.Combine(AppDir, "SystemConfig.sqlite");

        public static string BasePath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB");
        public static string AttachmentBasePath { get; private set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");

        private static readonly object _syncLock = new object();

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
                using (var cmd = new SQLiteCommand("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;", conn)) { cmd.ExecuteNonQuery(); }

                using (var cmd = new SQLiteCommand(conn))
                {
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS SysSettings (KeyName TEXT PRIMARY KEY, KeyValue TEXT);"; cmd.ExecuteNonQuery();
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS TableKeys (DbName TEXT, TableName TEXT, Col1 TEXT, Col2 TEXT, Col3 TEXT, Col4 TEXT, PRIMARY KEY(DbName, TableName));"; cmd.ExecuteNonQuery();
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS GridConfigs (DbName TEXT, TableName TEXT, ConfigType TEXT, ColName TEXT, ColValue TEXT, PRIMARY KEY(DbName, TableName, ConfigType, ColName));"; cmd.ExecuteNonQuery();
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS CustomStats (Module TEXT, StatName TEXT, Unit TEXT, Formula TEXT, PRIMARY KEY(Module, StatName));"; cmd.ExecuteNonQuery();
                    cmd.CommandText = @"CREATE TABLE IF NOT EXISTS SyncRules (Id INTEGER PRIMARY KEY AUTOINCREMENT, SrcDb TEXT, SrcTable TEXT, SrcMatchCol TEXT, SrcSyncCol TEXT, TgtDb TEXT, TgtTable TEXT, TgtMatchCol TEXT, TgtSyncCol TEXT, SyncType TEXT DEFAULT '單向同步');"; cmd.ExecuteNonQuery();
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS CustomMenus (Id INTEGER PRIMARY KEY AUTOINCREMENT, [分類] TEXT, [資料庫名] TEXT, [資料表名] TEXT);"; cmd.ExecuteNonQuery();
                    cmd.CommandText = @"CREATE TABLE IF NOT EXISTS DropdownConfigs (Id INTEGER PRIMARY KEY AUTOINCREMENT, TableName TEXT, ColName TEXT, ParentColName TEXT, ParentValue TEXT, Options TEXT, UNIQUE(TableName, ColName, ParentColName, ParentValue));"; cmd.ExecuteNonQuery();
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS AppLinks (Id INTEGER PRIMARY KEY AUTOINCREMENT, [選單名稱] TEXT, [執行路徑] TEXT);"; cmd.ExecuteNonQuery();
                    cmd.CommandText = @"CREATE TABLE IF NOT EXISTS System_DeleteLogs (Id INTEGER PRIMARY KEY AUTOINCREMENT, DbName TEXT, TableName TEXT, RecordId INTEGER, DeletedBy TEXT, DeletedTime TEXT);"; cmd.ExecuteNonQuery();

                    // 確保 v2 表結構存在
                    cmd.CommandText = @"CREATE TABLE IF NOT EXISTS ColumnFormulas (Id INTEGER PRIMARY KEY AUTOINCREMENT, DbName TEXT, TableName TEXT, TargetCol TEXT, MatchCol TEXT, StartDate TEXT, EndDate TEXT, Formula TEXT);"; 
                    cmd.ExecuteNonQuery();

                    // 🟢 動態升級表結構：新增 FormulaType
                    var cols = new List<string>();
                    cmd.CommandText = "PRAGMA table_info([ColumnFormulas])";
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) cols.Add(reader["name"].ToString());
                    }
                    if (!cols.Contains("FormulaType")) {
                        cmd.CommandText = "ALTER TABLE [ColumnFormulas] ADD COLUMN [FormulaType] TEXT DEFAULT '數學運算';";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        public static double GetUnitPrice(string category, DateTime date)
        {
            double price = 0;
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT UnitPrice FROM WaterPrices WHERE Category=@C AND StartDate <= @D AND EndDate >= @D ORDER BY StartDate DESC LIMIT 1", conn)) {
                        cmd.Parameters.AddWithValue("@C", category);
                        cmd.Parameters.AddWithValue("@D", date.ToString("yyyy-MM-dd"));
                        var res = cmd.ExecuteScalar();
                        if (res != null && res != DBNull.Value) return Convert.ToDouble(res);
                    }
                    using (var cmd = new SQLiteCommand("SELECT UnitPrice FROM WaterPrices WHERE Category=@C ORDER BY EndDate DESC LIMIT 1", conn)) {
                        cmd.Parameters.AddWithValue("@C", category);
                        var res = cmd.ExecuteScalar();
                        if (res != null && res != DBNull.Value) return Convert.ToDouble(res);
                    }
                }
            } catch { }
            return price;
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

            string attachPath = GetSysSetting("AttachmentBasePath", "");
            if (!string.IsNullOrEmpty(attachPath)) AttachmentBasePath = attachPath;
            if (!Directory.Exists(AttachmentBasePath)) Directory.CreateDirectory(AttachmentBasePath);
        }

        public static void SetBasePath(string newPath)
        {
            BasePath = newPath;
            SetSysSetting("BasePath", newPath);
            if (!Directory.Exists(BasePath)) Directory.CreateDirectory(BasePath);
        }

        public static void SetAttachmentBasePath(string newPath)
        {
            AttachmentBasePath = newPath;
            SetSysSetting("AttachmentBasePath", newPath);
            if (!Directory.Exists(AttachmentBasePath)) Directory.CreateDirectory(AttachmentBasePath);
        }

        // ========================================================
        // 🟢 欄位自訂公式 API (含 FormulaType)
        // ========================================================
        public static List<ColumnFormulaDef> GetTableFormulas(string dbName, string tableName)
        {
            var list = new List<ColumnFormulaDef>();
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT Id, TargetCol, MatchCol, StartDate, EndDate, FormulaType, Formula FROM ColumnFormulas WHERE DbName=@DB AND TableName=@TB", conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName); 
                        cmd.Parameters.AddWithValue("@TB", tableName);
                        using (var reader = cmd.ExecuteReader()) {
                            while (reader.Read()) { 
                                list.Add(new ColumnFormulaDef {
                                    Id = Convert.ToInt32(reader["Id"]),
                                    TargetCol = reader["TargetCol"].ToString(),
                                    MatchCol = reader["MatchCol"].ToString(),
                                    StartDate = reader["StartDate"].ToString(),
                                    EndDate = reader["EndDate"].ToString(),
                                    FormulaType = reader["FormulaType"].ToString() == "" ? "數學運算" : reader["FormulaType"].ToString(),
                                    Formula = reader["Formula"].ToString()
                                });
                            }
                        }
                    }
                }
            } catch { }
            return list;
        }

        public static void DeleteTableFormula(int id)
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("DELETE FROM ColumnFormulas WHERE Id=@Id", conn)) {
                        cmd.Parameters.AddWithValue("@Id", id);
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }
        }

        // ========================================================
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

        private static void EnsureNetworkConnection()
        {
            if (!Directory.Exists(BasePath))
            {
                throw new IOException($"網路連線不穩或實體路徑不可用！\n路徑：{BasePath}\n為保護資料庫免於損壞，系統已自動攔截此次儲存動作。請確認區網連線正常後再試一次。");
            }
        }

        private static void ExecuteWithRetry(string dbName, Action<SQLiteConnection> dbAction)
        {
            EnsureNetworkConnection(); 

            int maxRetries = 10;
            for (int i = 1; i <= maxRetries; i++) {
                try {
                    using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                        conn.Open(); 
                        using (var pragmaCmd = new SQLiteCommand("PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA busy_timeout=5000;", conn)) { pragmaCmd.ExecuteNonQuery(); }
                        dbAction(conn); 
                        return;
                    }
                } catch (SQLiteException ex) when (ex.ResultCode == SQLiteErrorCode.Locked || ex.ResultCode == SQLiteErrorCode.Busy) {
                    if (i == maxRetries) throw;
                    Thread.Sleep(100 * i);
                }
            }
        }

        private static void EnsureAuditColumns(SQLiteConnection conn, string tableName)
        {
            List<string> existingCols = new List<string>();
            using (var cmd = new SQLiteCommand($"PRAGMA table_info([{tableName}])", conn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read()) existingCols.Add(reader["name"].ToString());
            }

            if (!existingCols.Contains("最後修改人")) {
                using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] ADD COLUMN [最後修改人] TEXT", conn)) cmd.ExecuteNonQuery();
            }
            if (!existingCols.Contains("修改時間")) {
                using (var cmd = new SQLiteCommand($"ALTER TABLE [{tableName}] ADD COLUMN [修改時間] TEXT", conn)) cmd.ExecuteNonQuery();
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

        public static bool BulkSaveTable(string dbName, string tableName, DataTable dt, IProgress<int> progInt = null, IProgress<string> progStr = null)
        {
            if (dt == null || dt.Rows.Count == 0) return true; 

            try {
                EnsureNetworkConnection();

                using (var conn = new SQLiteConnection(GetConnString(dbName))) {
                    conn.Open();
                    using (var cmdPragma = new SQLiteCommand("PRAGMA journal_mode=WAL; PRAGMA busy_timeout=5000;", conn)) { cmdPragma.ExecuteNonQuery(); }

                    EnsureAuditColumns(conn, tableName);

                    string currentUser = Environment.UserName.Trim();
                    string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    using (var trans = conn.BeginTransaction()) {
                        var keys = GetTableKeys(dbName, tableName);

                        List<string> activeKeys = new List<string>();
                        if (!string.IsNullOrEmpty(keys.col1) && dt.Columns.Contains(keys.col1)) activeKeys.Add(keys.col1);
                        if (!string.IsNullOrEmpty(keys.col2) && dt.Columns.Contains(keys.col2)) activeKeys.Add(keys.col2);
                        if (!string.IsNullOrEmpty(keys.col3) && dt.Columns.Contains(keys.col3)) activeKeys.Add(keys.col3);
                        if (!string.IsNullOrEmpty(keys.col4) && dt.Columns.Contains(keys.col4)) activeKeys.Add(keys.col4);

                        int totalRows = dt.Rows.Count;
                        int currentRow = 0;

                        foreach (DataRow row in dt.Rows) {
                            currentRow++;
                            if (progInt != null && progStr != null && (currentRow % 50 == 0 || currentRow == totalRows)) {
                                progInt.Report((int)((double)currentRow / totalRows * 100));
                                progStr.Report($"正在高速寫入資料庫： 第 {currentRow} 筆 / 共 {totalRows} 筆");
                            }

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
                                    string safeKey = k.Replace(" ", "_").Replace("[", "").Replace("]", "").Replace("(", "_").Replace(")", "_").Replace("/", "_").Replace("-", "_");
                                    whereClauses.Add($"IFNULL([{k}], '') = IFNULL(@{safeKey}, '')");
                                }

                                string qCheck = $"SELECT Id FROM [{tableName}] WHERE " + string.Join(" AND ", whereClauses);

                                using (var cmdCheck = new SQLiteCommand(qCheck, conn, trans)) {
                                    foreach (var k in activeKeys) {
                                        string safeKey = k.Replace(" ", "_").Replace("[", "").Replace("]", "").Replace("(", "_").Replace(")", "_").Replace("/", "_").Replace("-", "_");
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
                                        if (col.ColumnName == "Id" || col.ColumnName == "最後修改人" || col.ColumnName == "修改時間") continue;
                                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "").Replace("(", "_").Replace(")", "_").Replace("/", "_").Replace("-", "_");
                                        sqlParts.Add($"[{col.ColumnName}]=@{safeParamName}");
                                        object val = row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value;
                                        cmd.Parameters.AddWithValue("@" + safeParamName, val);
                                    }
                                    sqlParts.Add("[最後修改人]=@SysUser");
                                    sqlParts.Add("[修改時間]=@SysTime");
                                    cmd.Parameters.AddWithValue("@SysUser", currentUser);
                                    cmd.Parameters.AddWithValue("@SysTime", currentTime);

                                    cmd.CommandText = sql + string.Join(", ", sqlParts) + " WHERE Id=" + targetId;
                                } else {
                                    List<string> colNames = new List<string>();
                                    List<string> paramNames = new List<string>();
                                    foreach (DataColumn col in dt.Columns) {
                                        if (col.ColumnName == "Id" || col.ColumnName == "最後修改人" || col.ColumnName == "修改時間") continue;
                                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "").Replace("(", "_").Replace(")", "_").Replace("/", "_").Replace("-", "_");
                                        colNames.Add($"[{col.ColumnName}]");
                                        paramNames.Add($"@{safeParamName}");
                                        object val = row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value;
                                        cmd.Parameters.AddWithValue("@" + safeParamName, val);
                                    }
                                    colNames.Add("[最後修改人]"); paramNames.Add("@SysUser");
                                    colNames.Add("[修改時間]"); paramNames.Add("@SysTime");
                                    cmd.Parameters.AddWithValue("@SysUser", currentUser);
                                    cmd.Parameters.AddWithValue("@SysTime", currentTime);

                                    cmd.CommandText = $"INSERT INTO [{tableName}] ({string.Join(", ", colNames)}) VALUES ({string.Join(", ", paramNames)})";
                                }
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit(); 
                    }
                }
                
                ThreadPool.QueueUserWorkItem(_ => RunSyncEngine(dbName, tableName));
                return true;
            } catch (Exception ex) {
                MessageBox.Show("儲存中斷：" + ex.Message, "系統保護機制", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        public static long UpsertRecord(string dbName, string tableName, DataRow row)
        {
            if (row.RowState == DataRowState.Deleted || row.RowState == DataRowState.Unchanged) return -1;

            long insertedId = -1;

            ExecuteWithRetry(dbName, conn => {
                EnsureAuditColumns(conn, tableName);

                bool isUpdate = row.Table.Columns.Contains("Id") && row["Id"] != DBNull.Value && Convert.ToInt32(row["Id"]) > 0;
                string currentUser = Environment.UserName.Trim();
                string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                var cmd = new SQLiteCommand(conn);
                if (isUpdate) {
                    List<string> sets = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id" || col.ColumnName == "最後修改人" || col.ColumnName == "修改時間") continue;
                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "").Replace("(", "_").Replace(")", "_").Replace("/", "_").Replace("-", "_");
                        sets.Add($"[{col.ColumnName}]=@{safeParamName}");
                        cmd.Parameters.AddWithValue("@" + safeParamName, row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value);
                    }
                    sets.Add("[最後修改人]=@SysUser");
                    sets.Add("[修改時間]=@SysTime");
                    cmd.Parameters.AddWithValue("@SysUser", currentUser);
                    cmd.Parameters.AddWithValue("@SysTime", currentTime);

                    cmd.CommandText = $"UPDATE [{tableName}] SET {string.Join(", ", sets)} WHERE Id=" + row["Id"];
                    cmd.ExecuteNonQuery();
                    
                    insertedId = Convert.ToInt64(row["Id"]);
                } else {
                    List<string> c = new List<string>();
                    List<string> v = new List<string>();
                    foreach (DataColumn col in row.Table.Columns) {
                        if (col.ColumnName == "Id" || col.ColumnName == "最後修改人" || col.ColumnName == "修改時間") continue;
                        string safeParamName = col.ColumnName.Replace(" ", "_").Replace("[", "").Replace("]", "").Replace("(", "_").Replace(")", "_").Replace("/", "_").Replace("-", "_");
                        c.Add($"[{col.ColumnName}]"); 
                        v.Add($"@{safeParamName}");
                        cmd.Parameters.AddWithValue("@" + safeParamName, row[col] != DBNull.Value ? (object)row[col].ToString().Trim() : DBNull.Value);
                    }
                    c.Add("[最後修改人]"); v.Add("@SysUser");
                    c.Add("[修改時間]"); v.Add("@SysTime");
                    cmd.Parameters.AddWithValue("@SysUser", currentUser);
                    cmd.Parameters.AddWithValue("@SysTime", currentTime);

                    cmd.CommandText = $"INSERT INTO [{tableName}] ({string.Join(", ", c)}) VALUES ({string.Join(", ", v)})";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "SELECT last_insert_rowid()";
                    object res = cmd.ExecuteScalar();
                    if (res != null && res != DBNull.Value)
                    {
                        insertedId = Convert.ToInt64(res);
                    }
                }
            });
            
            ThreadPool.QueueUserWorkItem(_ => RunSyncEngine(dbName, tableName));
            return insertedId;
        }

        public static void DeleteRecord(string dbName, string tableName, int id) 
        {
            try {
                using (var conn = new SQLiteConnection($"Data Source={SysConfigDbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("INSERT INTO System_DeleteLogs (DbName, TableName, RecordId, DeletedBy, DeletedTime) VALUES (@DB, @TB, @RId, @User, @Time)", conn)) {
                        cmd.Parameters.AddWithValue("@DB", dbName);
                        cmd.Parameters.AddWithValue("@TB", tableName);
                        cmd.Parameters.AddWithValue("@RId", id);
                        cmd.Parameters.AddWithValue("@User", Environment.UserName.Trim());
                        cmd.Parameters.AddWithValue("@Time", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        cmd.ExecuteNonQuery();
                    }
                }
            } catch { }

            ExecuteWithRetry(dbName, conn => {
                using (var cmd = new SQLiteCommand($"DELETE FROM [{tableName}] WHERE Id=@Id", conn)) { 
                    cmd.Parameters.AddWithValue("@Id", id); 
                    cmd.ExecuteNonQuery(); 
                }
            });
            ThreadPool.QueueUserWorkItem(_ => RunSyncEngine(dbName, tableName));
        }

        public static void RunSyncEngine(string triggerDb, string triggerTable)
        {
            lock (_syncLock)
            {
                try
                {
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
            }
        }

        public static bool IsAttachmentInUse(string dbName, string tableName, string relativePath)
        {
            bool inUse = false;
            try {
                ExecuteWithRetry(dbName, conn => {
                    string sql = $"SELECT COUNT(1) FROM [{tableName}] WHERE [附件檔案] LIKE @Path";
                    using (var cmd = new SQLiteCommand(sql, conn)) {
                        cmd.Parameters.AddWithValue("@Path", $"%{relativePath}%");
                        var res = cmd.ExecuteScalar();
                        if (res != null && Convert.ToInt32(res) > 0) {
                            inUse = true;
                        }
                    }
                });
            } catch { }
            return inUse;
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
