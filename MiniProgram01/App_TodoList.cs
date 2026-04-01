using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Threading;

public class App_TodoList : UserControl {
    private string activeFile;
    private string addedLog;
    private string doneLog;

    public App_TodoList TargetList { get; set; }
    private string moveBtnText;

    private TextBox inputField;
    private FlowLayoutPanel taskContainer;
    private Dictionary<string, Tuple<DateTime, string>> taskData = new Dictionary<string, Tuple<DateTime, string>>();
    
    private int dragInsertIndex = -1; 
    private static Color AppleBlue = Color.FromArgb(0, 122, 255);
    private static Font MainFont = new Font("Microsoft JhengHei UI", 10f);
    private MainForm mainForm;

    private readonly string[] colorCycle = { "Black", "Red", "DodgerBlue", "MediumOrchid", "DarkGreen", "DarkOrange" };

    public App_TodoList(MainForm parent, string filePrefix, string moveText) {
        this.mainForm = parent; 
        this.moveBtnText = moveText;
        this.Dock = DockStyle.Fill; // 【修正1】保證佈局填滿，不再縮成一團
        
        activeFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePrefix + "_active.txt");
        addedLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePrefix + "_history_added.txt");
        doneLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePrefix + "_history_completed.txt");

        this.BackColor = Color.FromArgb(245, 245, 247);
        this.Padding = new Padding(10);

        Panel top = new Panel() { Dock = DockStyle.Top, Height = 40
