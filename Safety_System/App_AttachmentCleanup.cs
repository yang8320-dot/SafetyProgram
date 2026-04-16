/// FILE: Safety_System/App_AttachmentCleanup.cs ///
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Safety_System
{
    public class App_AttachmentCleanup
    {
        private Button _btnScan;
        private Button _btnClean;
        private RichTextBox _rtbLog;
        private Label _lblStatus;

        // 儲存掃描出的多餘檔案絕對路徑
        private List<string> _orphanFiles = new List<string>();

        public Control GetView()
        {
            TableLayoutPanel main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                RowCount = 3,
                ColumnCount = 1
            };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // 1. 頂部說明與操作區
            GroupBox boxTop = new GroupBox { Text = "🧹 附件檔案空間深度清理", Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold), AutoSize = true, Padding = new Padding(15) };
            FlowLayoutPanel flpTop = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };

            Label lblDesc = new Label
            {
                Text = "此功能將掃描所有資料庫與資料表，並比對「附件」資料夾中的實體檔案。\n找出【未被任何資料紀錄綁定】的多餘孤兒檔案，幫助您釋放伺服器/硬碟空間。",
                AutoSize = true,
                Font = new Font("Microsoft JhengHei UI", 11F),
                ForeColor = Color.DimGray,
                Margin = new Padding(0, 0, 0, 15)
            };

            _btnScan = new Button { Text = "🔍 1. 開始全面掃描", Size = new Size(200, 45), BackColor = Color.SteelBlue, ForeColor = Color.White, Cursor = Cursors.Hand };
            _btnScan.Click += async (s, e) => await ScanAsync();

            _btnClean = new Button { Text = "🗑️ 2. 清除多餘檔案", Size = new Size(200, 45), BackColor = Color.IndianRed, ForeColor = Color.White, Cursor = Cursors.Hand, Enabled = false, Margin = new Padding(15, 0, 0, 0) };
            _btnClean.Click += async (s, e) => await CleanAsync();

            flpTop.Controls.Add(lblDesc);
            flpTop.SetFlowBreak(lblDesc, true);
            flpTop.Controls.Add(_btnScan);
            flpTop.Controls.Add(_btnClean);

            boxTop.Controls.Add(flpTop);

            // 2. 日誌顯示區
            _rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 11F),
                BackColor = Color.FromArgb(240, 240, 240),
                ReadOnly = true,
                Margin = new Padding(0, 15, 0, 15)
            };

            // 3. 底部狀態列
            _lblStatus = new Label
            {
                Text = "準備就緒，請點擊「開始全面掃描」。",
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue,
                AutoSize = true
            };

            main.Controls.Add(boxTop, 0, 0);
            main.Controls.Add(_rtbLog, 0, 1);
            main.Controls.Add(_lblStatus, 0, 2);

            return main;
        }

        private void Log(string message, Color? color = null)
        {
            if (_rtbLog.InvokeRequired)
            {
                _rtbLog.Invoke(new Action(() => Log(message, color)));
                return;
            }

            _rtbLog.SelectionStart = _rtbLog.TextLength;
            _rtbLog.SelectionLength = 0;
            _rtbLog.SelectionColor = color ?? Color.Black;
            _rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            _rtbLog.ScrollToCaret();
        }

        private async Task ScanAsync()
        {
            _btnScan.Enabled = false;
            _btnClean.Enabled = false;
            _rtbLog.Clear();
            _orphanFiles.Clear();
            _lblStatus.Text = "掃描進行中，請稍候...";
            _lblStatus.ForeColor = Color.Orange;

            Log("🚀 開始全面掃描...", Color.Blue);

            await Task.Run(() =>
            {
                // 1. 取得資料庫中所有的附件參照路徑
                HashSet<string> dbReferencedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                try
                {
                    string[] dbFiles = Directory.GetFiles(DataManager.BasePath, "*.sqlite");
                    Log($"📂 找到 {dbFiles.Length} 個資料庫檔案。");

                    foreach (var dbFile in dbFiles)
                    {
                        using (var conn = new SQLiteConnection($"Data Source={dbFile};Version=3;"))
                        {
                            conn.Open();
                            List<string> tables = new List<string>();

                            // 取得該 DB 所有的資料表
                            using (var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'", conn))
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read()) tables.Add(reader.GetString(0));
                            }

                            // 逐一資料表檢查是否含有 [附件檔案] 欄位
                            foreach (var table in tables)
                            {
                                bool hasAttachCol = false;
                                using (var cmd = new SQLiteCommand($"PRAGMA table_info([{table}])", conn))
                                using (var reader = cmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        if (reader["name"].ToString() == "附件檔案")
                                        {
                                            hasAttachCol = true;
                                            break;
                                        }
                                    }
                                }

                                // 若有附件欄位，則撈取所有非空的字串
                                if (hasAttachCol)
                                {
                                    using (var cmd = new SQLiteCommand($"SELECT [附件檔案] FROM [{table}] WHERE [附件檔案] IS NOT NULL AND [附件檔案] != ''", conn))
                                    using (var reader = cmd.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            string pathsStr = reader.GetString(0);
                                            string[] paths = pathsStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                            foreach (var p in paths)
                                            {
                                                // 統一轉為正斜線 / 以利比對
                                                dbReferencedPaths.Add(p.Replace("\\", "/").Trim());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    Log($"✅ 資料庫掃描完成，共有 {dbReferencedPaths.Count} 筆有效的檔案關聯。");
                }
                catch (Exception ex)
                {
                    Log($"❌ 資料庫掃描發生錯誤：{ex.Message}", Color.Red);
                    return;
                }

                // 2. 掃描實體附件資料夾
                try
                {
                    string attachRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");
                    if (!Directory.Exists(attachRootDir))
                    {
                        Log("⚠️ 找不到「附件」資料夾，目前無任何實體檔案。", Color.Orange);
                        return;
                    }

                    string[] physicalFiles = Directory.GetFiles(attachRootDir, "*.*", SearchOption.AllDirectories);
                    Log($"📂 實體資料夾掃描完成，共有 {physicalFiles.Length} 個檔案。");

                    int matchCount = 0;
                    foreach (var file in physicalFiles)
                    {
                        // 將絕對路徑轉為相對路徑 (例如: 附件/Water/WaterMeterReadings/2023-10/123.pdf)
                        string relPath = file.Substring(AppDomain.CurrentDomain.BaseDirectory.Length).TrimStart('\\', '/').Replace("\\", "/");

                        if (dbReferencedPaths.Contains(relPath))
                        {
                            matchCount++;
                        }
                        else
                        {
                            _orphanFiles.Add(file);
                            Log($"🔍 發現多餘檔案: {relPath}", Color.DimGray);
                        }
                    }

                    Log($"✅ 實體檔案比對完成！正常使用中的檔案: {matchCount} 個，多餘檔案: {_orphanFiles.Count} 個。", Color.Green);
                }
                catch (Exception ex)
                {
                    Log($"❌ 實體檔案掃描發生錯誤：{ex.Message}", Color.Red);
                }
            });

            _lblStatus.Text = $"掃描結束。發現 {_orphanFiles.Count} 個多餘檔案。";
            _lblStatus.ForeColor = _orphanFiles.Count > 0 ? Color.IndianRed : Color.ForestGreen;
            
            _btnScan.Enabled = true;
            _btnClean.Enabled = _orphanFiles.Count > 0;
        }

        private async Task CleanAsync()
        {
            if (_orphanFiles.Count == 0) return;

            if (MessageBox.Show($"確定要永久刪除這 {_orphanFiles.Count} 個多餘檔案嗎？\n刪除後無法復原！", "確認清理", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            _btnScan.Enabled = false;
            _btnClean.Enabled = false;
            _lblStatus.Text = "檔案清理中，請稍候...";
            _lblStatus.ForeColor = Color.Orange;

            Log("\n🗑️ 開始執行清理作業...", Color.Red);

            await Task.Run(() =>
            {
                int successCount = 0;
                foreach (var file in _orphanFiles)
                {
                    try
                    {
                        if (File.Exists(file))
                        {
                            File.Delete(file);
                            successCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"❌ 刪除失敗: {Path.GetFileName(file)} - {ex.Message}", Color.Red);
                    }
                }

                Log($"✅ 成功刪除 {successCount} 個多餘檔案！", Color.Green);

                // 清理空資料夾
                try
                {
                    string attachRootDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "附件");
                    if (Directory.Exists(attachRootDir))
                    {
                        int dirDeleted = CleanEmptyDirectories(attachRootDir);
                        if (dirDeleted > 0) Log($"🧹 同步清理了 {dirDeleted} 個空資料夾。");
                    }
                }
                catch { }
            });

            _orphanFiles.Clear();
            _lblStatus.Text = "清理作業完成！";
            _lblStatus.ForeColor = Color.ForestGreen;
            _btnScan.Enabled = true;
        }

        // 遞迴刪除空資料夾
        private int CleanEmptyDirectories(string startLocation)
        {
            int deletedCount = 0;
            foreach (var directory in Directory.GetDirectories(startLocation))
            {
                deletedCount += CleanEmptyDirectories(directory);
                if (Directory.GetFiles(directory).Length == 0 && Directory.GetDirectories(directory).Length == 0)
                {
                    try
                    {
                        Directory.Delete(directory, false);
                        deletedCount++;
                    }
                    catch { }
                }
            }
            return deletedCount;
        }
    }
}
