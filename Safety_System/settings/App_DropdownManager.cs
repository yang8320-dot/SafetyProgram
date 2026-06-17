/// FILE: Safety_System/settings/App_DropdownManager.cs ///
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Safety_System
{
    public class DropdownItemDef
    {
        public string Text { get; set; }
        public string IconBase64 { get; set; }

        public Image GetImage()
        {
            if (string.IsNullOrEmpty(IconBase64)) return null;
            try
            {
                byte[] bytes = Convert.FromBase64String(IconBase64);
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    return Image.FromStream(ms);
                }
            }
            catch { return null; }
        }
    }

    public class ReferenceDef 
    {
        public string SourceDb { get; set; }
        public string SourceTb { get; set; }
        public string SourceCol { get; set; }
    }

    public partial class App_DropdownManager : Form
    {
        private TabControl _tabControl;
        
        private class ItemMap 
        {
            public string EnName;
            public string ChName;
            public override string ToString() => string.IsNullOrEmpty(ChName) ? " " : ChName; 
        }

        private readonly Dictionary<string, (string ChDbName, Dictionary<string, string> Tables)> _dbMap;

        public static Dictionary<string, List<DropdownItemDef>> DropdownCache = new Dictionary<string, List<DropdownItemDef>>();
        public static Dictionary<string, List<DropdownItemDef>> MultiSelectCache = new Dictionary<string, List<DropdownItemDef>>();
        public static Dictionary<string, ReferenceDef> ReferenceCache = new Dictionary<string, ReferenceDef>();

        // 拖曳排版的狀態變數 (共用)
        private int _dragFromRowIndex = -1;
        private Rectangle _dragBox = Rectangle.Empty;

        public App_DropdownManager()
        {
            try 
            {
                string sqlMulti = "CREATE TABLE IF NOT EXISTS [MultiSelectConfigs] (Id INTEGER PRIMARY KEY AUTOINCREMENT, TableName TEXT, ColName TEXT, Options TEXT, UNIQUE(TableName, ColName));";
                DataManager.InitTable("SystemConfig", "MultiSelectConfigs", sqlMulti);

                string sqlRef = "CREATE TABLE IF NOT EXISTS [ReferenceDropdownConfigs] (Id INTEGER PRIMARY KEY AUTOINCREMENT, TargetDb TEXT, TargetTb TEXT, TargetCol TEXT, SourceDb TEXT, SourceTb TEXT, SourceCol TEXT, UNIQUE(TargetDb, TargetTb, TargetCol));";
                DataManager.InitTable("SystemConfig", "ReferenceDropdownConfigs", sqlRef);

                _dbMap = App_DbConfig.GetDbMapCache();
                RefreshConfiguredCache();
                InitializeComponent();
                
                LoadDropdownConfigs();
                LoadMultiSelectConfigs();
                LoadReferenceConfigs();

                RefreshMultiConfiguredList();
                RefreshRefConfiguredList();
            } 
            catch (Exception ex) 
            {
                MessageBox.Show($"初始化連動選單管理介面時發生嚴重錯誤：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "下拉選單與組合文字管理中心";
            this.Size = new Size(1650, 900);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.BackColor = Color.WhiteSmoke;
            this.Font = new Font("Microsoft JhengHei UI", 12F);

            _tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Microsoft JhengHei UI", 12F), Padding = new Point(20, 10) };

            TabPage tabSingle = new TabPage("一、單選下拉與多層連動設定");
            tabSingle.BackColor = Color.WhiteSmoke;
            BuildTabSingle(tabSingle);

            TabPage tabMulti = new TabPage("二、組合文字 (複選) 設定");
            tabMulti.BackColor = Color.WhiteSmoke;
            BuildTabMulti(tabMulti);

            TabPage tabRef = new TabPage("三、跨表參照下拉選單設定");
            tabRef.BackColor = Color.WhiteSmoke;
            BuildTabReference(tabRef);

            _tabControl.TabPages.Add(tabSingle);
            _tabControl.TabPages.Add(tabMulti);
            _tabControl.TabPages.Add(tabRef);
            
            this.Controls.Add(_tabControl);
        }

        // ========================================================
        // 共用方法區
        // ========================================================
        private List<string> GetColumnsSafe(string dbName, string tbName)
        {
            var cols = DataManager.GetColumnNames(dbName, tbName);
            if (cols == null || cols.Count <= 1) 
            {
                cols = new List<string>();
                if (TableSchemaManager.SchemaMap.ContainsKey(tbName)) 
                {
                    string schema = TableSchemaManager.SchemaMap[tbName];
                    var parts = schema.Split(',');
                    cols.Add("Id"); 
                    foreach(var p in parts) 
                    {
                        int start = p.IndexOf('[');
                        int end = p.IndexOf(']');
                        if (start >= 0 && end > start) 
                        {
                            cols.Add(p.Substring(start + 1, end - start - 1));
                        }
                    }
                }
            }
            return cols;
        }

        private bool IsColumnInDropdownCache(string tbName, string colName)
        {
            string prefix = $"{tbName}|{colName}|";
            foreach (var key in DropdownCache.Keys) {
                if (key.StartsWith(prefix)) return true;
            }
            return false;
        }

        private string CheckColumnConflict(string dbName, string tbName, string colName, string currentTab)
        {
            if (currentTab != "TabSingle" && IsColumnInDropdownCache(tbName, colName)) 
                return "「一、單選下拉與多層連動」";
                
            if (currentTab != "TabMulti" && MultiSelectCache.ContainsKey($"{tbName}|{colName}")) 
                return "「二、組合文字 (複選)」";
                
            if (currentTab != "TabRef" && ReferenceCache.ContainsKey($"{dbName}|{tbName}|{colName}")) 
                return "「三、跨表參照下拉選單」";

            return null;
        }

        // 共用的 DGV 拖曳事件
        private void DgvOptions_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            var hit = dgv.HitTest(e.X, e.Y);
            
            if (hit.RowIndex >= 0 && hit.ColumnIndex == -1 && !dgv.Rows[hit.RowIndex].IsNewRow)
            {
                _dragFromRowIndex = hit.RowIndex;
                Size dragSize = SystemInformation.DragSize;
                _dragBox = new Rectangle(new Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize);
            }
            else
            {
                _dragBox = Rectangle.Empty;
            }
        }

        private void DgvOptions_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                if (_dragBox != Rectangle.Empty && !_dragBox.Contains(e.X, e.Y))
                {
                    DataGridView dgv = (DataGridView)sender;
                    dgv.DoDragDrop(dgv.Rows[_dragFromRowIndex], DragDropEffects.Move);
                }
            }
        }

        private void DgvOptions_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void DgvOptions_DragDrop(object sender, DragEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            Point clientPoint = dgv.PointToClient(new Point(e.X, e.Y));
            var hit = dgv.HitTest(clientPoint.X, clientPoint.Y);
            
            int targetRowIndex = hit.RowIndex;

            if (targetRowIndex < 0 || targetRowIndex >= dgv.Rows.Count - 1) 
            {
                targetRowIndex = dgv.Rows.Count - 2; 
            }

            if (e.Data.GetDataPresent(typeof(DataGridViewRow)))
            {
                DataGridViewRow dragRow = (DataGridViewRow)e.Data.GetData(typeof(DataGridViewRow));
                int sourceRowIndex = dragRow.Index;

                if (sourceRowIndex != targetRowIndex && targetRowIndex >= 0 && sourceRowIndex >= 0)
                {
                    dgv.EndEdit(); 
                    dgv.Rows.RemoveAt(sourceRowIndex);
                    dgv.Rows.Insert(targetRowIndex, dragRow);
                    dgv.ClearSelection();
                    dgv.CurrentCell = dgv[0, targetRowIndex];
                }
            }
        }

        private void MoveRowUp(DataGridView dgv)
        {
            if (dgv.CurrentCell != null)
            {
                int idx = dgv.CurrentCell.RowIndex;
                int colIdx = dgv.CurrentCell.ColumnIndex;
                if (idx > 0 && !dgv.Rows[idx].IsNewRow)
                {
                    dgv.EndEdit();
                    DataGridViewRow row = dgv.Rows[idx];
                    dgv.Rows.RemoveAt(idx);
                    dgv.Rows.Insert(idx - 1, row);
                    dgv.CurrentCell = dgv[colIdx, idx - 1];
                }
            }
        }

        private void MoveRowDown(DataGridView dgv)
        {
            if (dgv.CurrentCell != null)
            {
                int idx = dgv.CurrentCell.RowIndex;
                int colIdx = dgv.CurrentCell.ColumnIndex;
                if (idx < dgv.Rows.Count - 2 && !dgv.Rows[idx].IsNewRow)
                {
                    dgv.EndEdit();
                    DataGridViewRow row = dgv.Rows[idx];
                    dgv.Rows.RemoveAt(idx);
                    dgv.Rows.Insert(idx + 1, row);
                    dgv.CurrentCell = dgv[colIdx, idx + 1];
                }
            }
        }
    }
}
