// DataManager.cs 擴充方法

/// <summary>
/// 智慧儲存：根據是否有 Id 決定是 新增 還是 複寫
/// </summary>
public static void UpsertWaterRecord(DataRow row)
{
    ExecuteWithRetry(conn => {
        bool isUpdate = row["Id"] != DBNull.Value;
        string sql;

        if (isUpdate)
        {
            // 複寫 (Update) 邏輯 - 動態組建 SQL 以支援動態欄位
            StringBuilder sb = new StringBuilder("UPDATE WaterMeterReadings SET ");
            foreach (DataColumn col in row.Table.Columns)
            {
                if (col.ColumnName == "Id") continue;
                sb.Append(string.Format("[{0}] = @{0}, ", col.ColumnName));
            }
            sb.Remove(sb.Length - 2, 2); // 移除最後的逗號
            sb.Append(" WHERE Id = @Id");
            sql = sb.ToString();
        }
        else
        {
            // 新增 (Insert) 邏輯
            StringBuilder sbCols = new StringBuilder("INSERT INTO WaterMeterReadings (");
            StringBuilder sbVals = new StringBuilder("VALUES (");
            foreach (DataColumn col in row.Table.Columns)
            {
                if (col.ColumnName == "Id") continue;
                sbCols.Append(string.Format("[{0}], ", col.ColumnName));
                sbVals.Append(string.Format("@{0}, ", col.ColumnName));
            }
            sbCols.Remove(sbCols.Length - 2, 2).Append(") ");
            sbVals.Remove(sbVals.Length - 2, 2).Append(") ");
            sql = sbCols.ToString() + sbVals.ToString();
        }

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            foreach (DataColumn col in row.Table.Columns)
            {
                cmd.Parameters.AddWithValue("@" + col.ColumnName, row[col.ColumnName] ?? "");
            }
            cmd.ExecuteNonQuery();
        }
    });
}
