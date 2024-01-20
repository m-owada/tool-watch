using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualBasic.FileIO;

[assembly: AssemblyVersion("2.0.0.0")]
[assembly: AssemblyFileVersion ("2.0.0.0")]
[assembly: AssemblyInformationalVersion("2.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("WATCH")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (c) 2024 m-owada.")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

class Program
{
    [STAThread]
    static void Main()
    {
        bool createdNew;
        var mutex = new Mutex(true, @"Global\WATCH_COPYRIGHT_2024_M-OWADA", out createdNew);
        try
        {
            if(!createdNew)
            {
                MessageBox.Show("複数起動はできません。");
                return;
            }
            Application.EnableVisualStyles();
            Application.ThreadException += (object sender, ThreadExceptionEventArgs e) =>
            {
                throw new Exception(e.Exception.Message);
            };
            var mainForm = new MainForm();
            Application.Run();
        }
        catch(Exception e)
        {
            MessageBox.Show(e.Message, e.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
        }
        finally
        {
            mutex.ReleaseMutex();
            mutex.Close();
        }
    }
}

class MainForm : Form
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    
    private System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
    private NotifyIcon notifyIcon = new NotifyIcon();
    
    private string directory = string.Empty;
    private string encoding = "shift_jis";
    private int history = 10;
    private bool doublebuffered = false;
    
    public MainForm()
    {
        // フォーム
        this.Text = "WATCH";
        this.ShowInTaskbar = false;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        
        // タイマー
        timer.Interval = 60000;
        timer.Enabled = true;
        timer.Tick += Timer_Tick;
        
        // 設定ファイル
        LoadConfigSetting();
        
        // メニュー
        var menu = new ContextMenuStrip();
        menu.Items.Add(SetMenuItem("&ログ"));
        menu.Items.Add(SetMenuItem("&集計", Sum_Click));
        menu.Items.Add(SetMenuItem("&ビューア", Viewer_Click));
        menu.Items.Add(SetMenuItem("&終了", Close_Click));
        
        // タスクトレイ
        notifyIcon.Icon = this.Icon;
        notifyIcon.Visible = true;
        notifyIcon.Text = this.Text;
        notifyIcon.ContextMenuStrip = menu;
        notifyIcon.DoubleClick += NotifyIcon_DoubleClick;
        
        // バルーンチップ
        notifyIcon.BalloonTipTitle = this.Text;
        notifyIcon.BalloonTipText = "アクティブウィンドウのプロセスを監視します。";
        notifyIcon.ShowBalloonTip(10000);
    }
    
    private ToolStripMenuItem SetMenuItem(string text)
    {
        var menuItem = new ToolStripMenuItem();
        menuItem.Text = text;
        GetLogName();
        SetMenuDropDownItems(menuItem);
        return menuItem;
    }
    
    private ToolStripMenuItem SetMenuItem(string text, EventHandler handler)
    {
        var menuItem = new ToolStripMenuItem();
        menuItem.Text = text;
        menuItem.Click += handler;
        return menuItem;
    }
    
    private void SetMenuDropDownItems(ToolStripMenuItem item)
    {
        var cnt = 0;
        item.DropDownItems.Clear();
        foreach(var file in Directory.EnumerateFiles(GetLogDirectory(), "*.log", System.IO.SearchOption.TopDirectoryOnly).OrderByDescending(f => f))
        {
            item.DropDownItems.Add(SetMenuItem(Path.GetFileName(file), Log_Click));
            cnt++;
            if(cnt >= history)
            {
                break;
            }
        }
    }
    
    private void Log_Click(object sender, EventArgs e)
    {
        var name1 = GetLogDirectory() + @"\" + ((ToolStripMenuItem)sender).Text;
        if(File.Exists(name1))
        {
            var name2 = "watch.txt";
            File.Copy(name1, name2, true);
            Process.Start(name2);
        }
        else
        {
            MessageBox.Show("選択されたログファイルは存在しません。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    
    private void Sum_Click(object sender, EventArgs e)
    {
        SetMenuItemsEnabled(false);
        SubForm1 subForm = new SubForm1(GetLogDirectory(), encoding);
        subForm.ShowDialog();
        SetMenuItemsEnabled(true);
    }
    
    private void Viewer_Click(object sender, EventArgs e)
    {
        SetMenuItemsEnabled(false);
        SubForm2 subForm = new SubForm2(GetLogDirectory(), encoding, doublebuffered);
        subForm.ShowDialog();
        SetMenuItemsEnabled(true);
    }
    
    private void Close_Click(object sender, EventArgs e)
    {
        if(MessageBox.Show("終了しますか？", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        {
            notifyIcon.Dispose();
            Application.Exit();
        }
    }
    
    private void NotifyIcon_DoubleClick(object sender, EventArgs e)
    {
        notifyIcon.ContextMenuStrip.Items[2].PerformClick();
    }
    
    private void SetMenuItemsEnabled(bool enabled)
    {
        notifyIcon.ContextMenuStrip.Items[0].Enabled = enabled;
        notifyIcon.ContextMenuStrip.Items[1].Enabled = enabled;
        notifyIcon.ContextMenuStrip.Items[2].Enabled = enabled;
    }
    
    private string GetLogName()
    {
        var name = GetLogDirectory() + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".log";
        if(!File.Exists(name))
        {
            var log = "\"日時\",\"インターバル\",\"ID\",\"プロセス\",\"タイトル\"" + Environment.NewLine;
            File.AppendAllText(name, log, Encoding.GetEncoding(encoding));
            if(notifyIcon.ContextMenuStrip != null)
            {
                SetMenuDropDownItems((ToolStripMenuItem)notifyIcon.ContextMenuStrip.Items[0]);
            }
        }
        return name;
    }
    
    private string GetLogDirectory()
    {
        var dir = Environment.CurrentDirectory;
        if(directory.Length > 0)
        {
            Directory.CreateDirectory(directory);
            dir = directory;
        }
        return dir;
    }
    
    private void Timer_Tick(object sender, EventArgs e)
    {
        int id;
        GetWindowThreadProcessId(GetForegroundWindow(), out id);
        var name = GetLogName();
        var log = "\"" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\","
                + "\"" + timer.Interval + "\","
                + "\"" + Process.GetProcessById(id).Id + "\","
                + "\"" + Process.GetProcessById(id).ProcessName + "\","
                + "\"" + Process.GetProcessById(id).MainWindowTitle + "\"" + Environment.NewLine;
        File.AppendAllText(name, log, Encoding.GetEncoding(encoding));
    }
    
    private void LoadConfigSetting()
    {
        var config = LoadXmlFile("config.xml").Element("config");
        timer.Interval = GetXmlValue(config, "interval", timer.Interval);
        directory = GetXmlValue(config, "directory", directory);
        encoding = GetXmlValue(config, "encoding", encoding);
        history = GetXmlValue(config, "history", history);
        doublebuffered = GetXmlValue(config, "doublebuffered", doublebuffered);
    }
    
    private int GetXmlValue(XElement node, string name, int init)
    {
        var val = init;
        if(node.Element(name) != null)
        {
            Int32.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private string GetXmlValue(XElement node, string name, string init)
    {
        var val = init;
        if(node.Element(name) != null)
        {
            val = node.Element(name).Value;
        }
        return val;
    }
    
    private bool GetXmlValue(XElement node, string name, bool init)
    {
        var val = init;
        if(node.Element(name) != null)
        {
            Boolean.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private XDocument LoadXmlFile(string file)
    {
        if(File.Exists(file))
        {
            return XDocument.Load(file);
        }
        else
        {
            var xml = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("config",
                    new XElement("interval", 60000),
                    new XElement("directory", "log"),
                    new XElement("encoding", "shift_jis"),
                    new XElement("history", 10),
                    new XElement("doublebuffered", false)
                )
            );
            xml.Save(file);
            return xml;
        }
    }
}

class SubForm1 : Form
{
    private DateTimePicker dtp1 = new DateTimePicker();
    private DateTimePicker dtp2 = new DateTimePicker();
    private Button button1 = new Button();
    private ToolStripStatusLabel toolStripStatusLabel = new ToolStripStatusLabel();
    private const string statusDefaultValue = "日付を入力して集計ボタンをクリックしてください。";
    
    private string directory = string.Empty;
    private string encoding = string.Empty;
    private CancellationTokenSource cts = new CancellationTokenSource();
    
    public SubForm1(string directory, string encoding)
    {
        // フォーム
        this.Text = "WATCH";
        this.Size = new Size(260, 120);
        this.MaximizeBox = false;
        this.MinimizeBox = true;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.FormClosing += OnFormClosing;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        
        // 日付入力1
        dtp1.Location = new Point(10, 10);
        dtp1.Size = new Size(105, 20);
        dtp1.Value = DateTime.Today.AddDays(-6);
        dtp1.ValueChanged += Dtp1_ValueChanged;
        this.Controls.Add(dtp1);
        
        // ～
        var label1 = new Label();
        label1.Location = new Point(120, 15);
        label1.Size = new Size(20, 20);
        label1.Text = "～";
        this.Controls.Add(label1);
        
        // 日付入力2
        dtp2.Location = new Point(140, 10);
        dtp2.Size = new Size(105, 20);
        dtp2.Value = DateTime.Today;
        dtp2.ValueChanged += Dtp2_ValueChanged;
        this.Controls.Add(dtp2);
        
        // 集計ボタン
        button1.Location = new Point(10, 40);
        button1.Size = new Size(235, 20);
        button1.Text = "集計";
        button1.Click += Button1_Click;
        this.Controls.Add(button1);
        
        // ステータスバー
        var statusStrip = new StatusStrip();
        statusStrip.SizingGrip = false;
        statusStrip.Items.Add(toolStripStatusLabel);
        toolStripStatusLabel.Text = statusDefaultValue;
        this.Controls.Add(statusStrip);
        
        this.directory = directory;
        this.encoding = encoding;
    }
    
    private void OnFormClosing(object sender, FormClosingEventArgs e)
    {
        cts.Cancel();
    }
    
    private void Dtp1_ValueChanged(object sender, EventArgs e)
    {
        if(dtp1.Value > dtp2.Value)
        {
            dtp2.Value = dtp1.Value;
        }
    }
    
    private void Dtp2_ValueChanged(object sender, EventArgs e)
    {
        if(dtp1.Value > dtp2.Value)
        {
            dtp1.Value = dtp2.Value;
        }
    }
    
    private void Button1_Click(object sender, EventArgs e)
    {
        SummaryExecute();
    }
    
    private async void SummaryExecute()
    {
        dtp1.Enabled = false;
        dtp2.Enabled = false;
        button1.Enabled = false;
        
        await Task.Run(() => SummaryLogFiles(cts.Token));
        
        dtp1.Enabled = true;
        dtp2.Enabled = true;
        button1.Enabled = true;
        toolStripStatusLabel.Text = statusDefaultValue;
    }
    
    private void SummaryLogFiles(CancellationToken ct)
    {
        var d1 = dtp1.Value.ToString("yyyyMMdd") + ".log";
        var d2 = dtp2.Value.ToString("yyyyMMdd") + ".log";
        
        Dictionary<Tuple<string, string>, long> table = new Dictionary<Tuple<string, string>, long>();
        
        foreach(var file in Directory.EnumerateFiles(directory, "*.log"))
        {
            var name = Path.GetFileName(file);
            if(string.Compare(d1, name) <= 0 && string.Compare(d2, name) >= 0)
            {
                using(var parser = new TextFieldParser(file, Encoding.GetEncoding(encoding)))
                {
                    toolStripStatusLabel.Text = "集計中...(" +  name + ")";
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    parser.HasFieldsEnclosedInQuotes = false;
                    parser.TrimWhiteSpace = true;
                    while(!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if(fields.Count() < 5)
                        {
                            continue;
                        }
                        for(var i = 0; i < fields.Length; i++)
                        {
                            fields[i] = fields[i].Trim('"');
                        }
                        long interval;
                        if(!long.TryParse(fields[1], out interval))
                        {
                            continue;
                        }
                        var key = Tuple.Create(fields[3], fields[4]);
                        if(table.ContainsKey(key))
                        {
                            table[key] += interval;
                        }
                        else
                        {
                            table.Add(key, interval);
                        }
                        if(ct.IsCancellationRequested)
                        {
                            return;
                        }
                    }
                }
            }
        }
        if(table.Count() > 0)
        {
            CreateList(table);
        }
        else
        {
            MessageBox.Show("指定された期間内に記録されたログファイルが存在しません。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
    
    private void CreateList(Dictionary<Tuple<string, string>, long> table)
    {
        var name = "watch.txt";
        if(File.Exists(name))
        {
            File.Delete(name);
        }
        
        // ヘッダ行
        var header = "\"プロセス\",\"タイトル\",\"時間\"" + Environment.NewLine;
        File.AppendAllText(name, header, Encoding.GetEncoding(encoding));
        
        // 明細行
        foreach(var row in table.OrderByDescending(r => r.Value).ThenBy(r => r.Key.Item1).ThenBy(r => r.Key.Item2))
        {
            var log = "\"" + row.Key.Item1 + "\",\"" + row.Key.Item2 + "\",\"" + TimeSpan.FromMilliseconds(row.Value) + "\"" + Environment.NewLine;
            File.AppendAllText(name, log, Encoding.GetEncoding(encoding));
        }
        Process.Start(name);
    }
}

class SubForm2 : Form
{
    private DateTimePicker dtp1 = new DateTimePicker();
    private Button button1 = new Button();
    private Button button2 = new Button();
    private Button button3 = new Button();
    private Button button4 = new Button();
    private TextBox textBox1 = new TextBox();
    private DataGridView dataGridView1 = new DataGridView();
    private SortableBindingList<Columns1> dataSource1 = new SortableBindingList<Columns1>();
    
    private string directory = string.Empty;
    private string encoding = string.Empty;
    private bool doublebuffered = false;
    
    public SubForm2(string directory, string encoding, bool doublebuffered)
    {
        // フォーム
        this.Text = "WATCH";
        this.Size = new Size(520, 490);
        this.MaximizeBox = true;
        this.MinimizeBox = true;
        this.MinimumSize = this.Size;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        this.Load += OnFormLoad;
        
        // ラベル1
        var label1 = new Label();
        label1.Location = new Point(10, 10);
        label1.Size = new Size(55, 20);
        label1.TextAlign = ContentAlignment.MiddleLeft;
        label1.Text = "対象日付";
        this.Controls.Add(label1);
        
        // 日付入力
        dtp1.Location = new Point(70, 10);
        dtp1.Size = new Size(140, 20);
        dtp1.Value = DateTime.Today;
        dtp1.Format = DateTimePickerFormat.Custom;
        dtp1.CustomFormat = "yyyy/MM/dd dddd";
        dtp1.ValueChanged += Dtp1_ValueChanged;
        this.Controls.Add(dtp1);
        
        // ボタン1
        button1.Location = new Point(220, 10);
        button1.Size = new Size(20, 20);
        button1.Text = "<";
        button1.Click += Button1_Click;
        this.Controls.Add(button1);
        
        // ボタン2
        button2.Location = new Point(245, 10);
        button2.Size = new Size(40, 20);
        button2.Text = "今日";
        button2.Click += Button2_Click;
        this.Controls.Add(button2);
        
        // ボタン3
        button3.Location = new Point(290, 10);
        button3.Size = new Size(20, 20);
        button3.Text = ">";
        button3.Click += Button3_Click;
        this.Controls.Add(button3);
        
        // ラベル2
        var label2 = new Label();
        label2.Location = new Point(320, 10);
        label2.Size = new Size(55, 20);
        label2.TextAlign = ContentAlignment.MiddleLeft;
        label2.Text = "合計時間";
        this.Controls.Add(label2);
        
        // テキストボックス1
        textBox1.Location = new Point(380, 10);
        textBox1.Size = new Size(60, 20);
        textBox1.Text = string.Empty;
        textBox1.TextAlign = HorizontalAlignment.Center;
        textBox1.ReadOnly = true;
        this.Controls.Add(textBox1);
        
        // ボタン4
        button4.Location = new Point(455, 10);
        button4.Size = new Size(40, 20);
        button4.Text = "ログ";
        button4.Click += Button4_Click;
        this.Controls.Add(button4);
        
        // データグリッドビュー1
        dataGridView1.Location = new Point(10, 40);
        dataGridView1.Size = new Size(485, 400);
        dataGridView1.DataSource = dataSource1;
        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        dataGridView1.AllowUserToAddRows = false;
        dataGridView1.AllowUserToDeleteRows = false;
        dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
        dataGridView1.AllowUserToResizeColumns = false;
        dataGridView1.AllowUserToResizeRows = false;
        dataGridView1.MultiSelect = false;
        dataGridView1.ReadOnly = true;
        dataGridView1.RowHeadersVisible = false;
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
        dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(dataGridView1);
        
        this.directory = directory;
        this.encoding = encoding;
        this.doublebuffered = doublebuffered;
    }
    
    private class Columns1
    {
        [DisplayName("開始日時")]
        public string ShowStartTime { get { return StartTime.ToString("yyyy/MM/dd HH:mm:ss"); } }
        
        [DisplayName("作業時間")]
        public TimeSpan WorkTime { get; set; }
        
        [DisplayName("プロセス")]
        public string ProcessName { get; set; }
        
        [DisplayName("タイトル")]
        public string Title { get; set; }
        
        [Browsable(false)]
        public DateTime StartTime { get; set; }
    }
    
    private void OnFormLoad(object sender, EventArgs e)
    {
        if(doublebuffered)
        {
            dataGridView1.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, dataGridView1, new object[]{ true });
        }
        SetDataGridView1();
    }
    
    private void Dtp1_ValueChanged(object sender, EventArgs e)
    {
        SetDataGridView1();
    }
    
    private void Button1_Click(object sender, EventArgs e)
    {
        dtp1.Value = dtp1.Value.AddDays(-1);
    }
    
    private void Button2_Click(object sender, EventArgs e)
    {
        if(dtp1.Value == DateTime.Today)
        {
            SetDataGridView1();
        }
        else
        {
            dtp1.Value = DateTime.Today;
        }
    }
    
    private void Button3_Click(object sender, EventArgs e)
    {
        dtp1.Value = dtp1.Value.AddDays(1);
    }
    
    private void Button4_Click(object sender, EventArgs e)
    {
        var name1 = directory + @"\" + dtp1.Value.ToString("yyyyMMdd") + ".log";
        if(File.Exists(name1))
        {
            var name2 = "watch.txt";
            File.Copy(name1, name2, true);
            Process.Start(name2);
        }
        else
        {
            MessageBox.Show("対象日付のログファイルは存在しません。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
    
    private void SetDataGridView1()
    {
        dataSource1.Clear();
        dataSource1.RemoveSort();
        dataSource1.RaiseListChangedEvents = false;
        var totalTime = new TimeSpan();
        var path = directory + @"\" + dtp1.Value.ToString("yyyyMMdd") + ".log";
        if(File.Exists(path))
        {
            using(var parser = new TextFieldParser(path, Encoding.GetEncoding(encoding)))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = false;
                parser.TrimWhiteSpace = true;
                while(!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    if(fields.Count() < 5)
                    {
                        continue;
                    }
                    for(var i = 0; i < fields.Length; i++)
                    {
                        fields[i] = fields[i].Trim('"');
                    }
                    DateTime dateTime;
                    if(!DateTime.TryParse(fields[0], out dateTime))
                    {
                        continue;
                    }
                    long interval;
                    if(!long.TryParse(fields[1], out interval))
                    {
                        continue;
                    }
                    totalTime += TimeSpan.FromMilliseconds(interval);
                    if(dataSource1.Count > 0)
                    {
                        if(dataSource1[dataSource1.Count - 1].ProcessName == fields[3] && dataSource1[dataSource1.Count - 1].Title == fields[4])
                        {
                            dataSource1[dataSource1.Count - 1].WorkTime += TimeSpan.FromMilliseconds(interval);
                            continue;
                        }
                    }
                    var row = new Columns1();
                    row.StartTime = dateTime;
                    row.WorkTime = TimeSpan.FromMilliseconds(interval);
                    row.ProcessName = fields[3];
                    row.Title = fields[4];
                    dataSource1.Add(row);
                }
            }
        }
        textBox1.Text = totalTime.ToString();
        dataSource1.RaiseListChangedEvents = true;
        dataSource1.ResetBindings();
        ResizeDataGridView1();
    }
    
    private void ResizeDataGridView1()
    {
        if(dataSource1.Count > 0)
        {
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        else
        {
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
    }
}

class SortableBindingList<T> : BindingList<T>
{
    private PropertyDescriptor sortProperty;
    private ListSortDirection sortDirection = ListSortDirection.Ascending;
    private bool isSorted;
    
    public SortableBindingList()
    {
    }
    
    public SortableBindingList(IList<T> list) : base(list)
    {
    }
    
    public void RemoveSort()
    {
        RemoveSortCore();
    }
    
    protected override bool SupportsSortingCore
    {
        get { return true; }
    }
    
    protected override PropertyDescriptor SortPropertyCore
    {
        get { return sortProperty; }
    }
    
    protected override ListSortDirection SortDirectionCore
    {
        get { return sortDirection; }
    }
    
    protected override bool IsSortedCore
    {
        get { return isSorted; }
    }
    
    protected override void RemoveSortCore()
    {
        sortProperty = null;
        sortDirection = ListSortDirection.Ascending;
        isSorted = false;
    }
    
    protected override void ApplySortCore(PropertyDescriptor property, ListSortDirection direction)
    {
        var list = Items as List<T>;
        if(list != null)
        {
            sortProperty = property;
            sortDirection = direction;
            isSorted = true;
            list.Sort(Compare);
            OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
        }
    }
    
    private int Compare(T x, T y)
    {
        var result = OnComparison(x, y);
        return (sortDirection == ListSortDirection.Ascending) ? result : - result;
    }
    
    private int OnComparison(T x, T y)
    {
        object xValue = (x == null) ? null : sortProperty.GetValue(x);
        object yValue = (y == null) ? null : sortProperty.GetValue(y);
        
        if(xValue == null)
        {
            return (yValue == null) ? 0 : -1;
        }
        if(yValue == null)
        {
            return 1;
        }
        if(xValue is IComparable)
        {
            return ((IComparable)xValue).CompareTo(yValue);
        }
        if(xValue.Equals(yValue))
        {
            return 0;
        }
        return xValue.ToString().CompareTo(yValue.ToString());
    }
}
