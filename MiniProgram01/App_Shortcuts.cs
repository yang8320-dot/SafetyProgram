using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public class App_Shortcuts : UserControl {
    private string shortcutsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "shortcuts.txt");
    private FlowLayoutPanel listPanel;
    private TextBox txtPath;
    private class ShortcutItem { public string Name, Path; }
    private List<ShortcutItem> items = new List<ShortcutItem>();

    public App_Shortcuts(MainForm parent) {
        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        FlowLayoutPanel top = new FlowLayoutPanel() { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown };
        txtPath = new TextBox() { Width = 310 };
        FlowLayoutPanel btns = new FlowLayoutPanel() { AutoSize = true };
        Button btnF = new Button() { Text = "選檔", Width = 60 };
        btnF.Click += (s, e) => { using(OpenFileDialog d=new OpenFileDialog()) if(d.ShowDialog()==DialogResult.OK) txtPath.Text=d.FileName; };
        Button btnD = new Button() { Text = "選資料夾", Width = 80 };
        btnD.Click += (s, e) => { using(FolderBrowserDialog d=new FolderBrowserDialog()) if(d.ShowDialog()==DialogResult.OK) txtPath.Text=d.SelectedPath; };
        Button btnA = new Button() { Text = "新增", Width = 60, BackColor = Color.FromArgb(0, 122, 255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnA.Click += (s, e) => Add(txtPath.Text);
        btns.Controls.AddRange(new Control[]{ btnF, btnD, btnA });
        top.Controls.AddRange(new Control[]{ txtPath, btns });

        listPanel = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.White };
        this.Controls.Add(listPanel); this.Controls.Add(top);
        LoadData();
    }

    private void Add(string p) {
        if (string.IsNullOrEmpty(p)) return;
        items.Add(new ShortcutItem() { Name = Path.GetFileName(p) ?? p, Path = p });
        SaveData(); RefreshUI(); txtPath.Text = "";
    }

    private void RefreshUI() {
        listPanel.Controls.Clear();
        foreach (var i in items) {
            Panel p = new Panel() { Width = 330, Height = 40, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 2, 0, 2) };
            Label l = new Label() { Text = i.Name, Location = new Point(5, 10), Width = 150, Cursor = Cursors.Hand };
            l.Click += (s, e) => { string n = ShowInputBox("改名:", "重新命名", i.Name); if(!string.IsNullOrEmpty(n)) { i.Name = n; SaveData(); RefreshUI(); } };
            Button bO = new Button() { Text = "開啟", Left = 160, Top = 5, Width = 60 };
            bO.Click += (s, e) => { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(i.Path){ UseShellExecute=true }); } catch { MessageBox.Show("失效路徑"); } };
            Button bD = new Button() { Text = "✕", Left = 225, Top = 5, Width = 30, ForeColor = Color.Red };
            bD.Click += (s, e) => { items.Remove(i); SaveData(); RefreshUI(); };
            p.Controls.AddRange(new Control[] { l, bO, bD }); listPanel.Controls.Add(p);
        }
    }

    private void SaveData() {
        List<string> l = new List<string>(); foreach(var i in items) l.Add(i.Name + "|" + i.Path);
        File.WriteAllLines(shortcutsFile, l);
    }

    private void LoadData() {
        if(!File.Exists(shortcutsFile)) return;
        foreach(string l in File.ReadAllLines(shortcutsFile)){ var p=l.Split('|'); if(p.Length>=2) items.Add(new ShortcutItem(){ Name=p[0], Path=p[1] }); }
        RefreshUI();
    }

    public static string ShowInputBox(string p, string t, string d) {
        Form f = new Form() { Width = 300, Height = 150, Text = t, StartPosition = FormStartPosition.CenterScreen };
        TextBox x = new TextBox() { Left = 20, Top = 40, Width = 240, Text = d };
        Button b = new Button() { Text = "OK", Left = 180, Top = 80, DialogResult = DialogResult.OK };
        f.Controls.AddRange(new Control[] { x, b }); f.AcceptButton = b;
        return f.ShowDialog() == DialogResult.OK ? x.Text : "";
    }
}
