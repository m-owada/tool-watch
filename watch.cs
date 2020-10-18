using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualBasic.FileIO;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion ("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("WATCH")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (c) 2020 m-owada.")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

class Program
{
    [STAThread]
    static void Main()
    {
        bool createdNew;
        var mutex = new Mutex(true, @"Global\WATCH_COPYRIGHT_2020_M-OWADA", out createdNew);
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
            MessageBox.Show(e.Message, e.Source);
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
    
    System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
    
    private string encoding = "shift_jis";
    private int history = 10;
    
    public MainForm()
    {
        // フォーム
        this.Text = "WATCH";
        this.ShowInTaskbar = false;
        
        // タイマー
        timer.Interval = 60000;
        timer.Enabled = true;
        timer.Tick += Timer_Tick;
        
        // 設定ファイル
        LoadConfigSetting();
        
        // メニュー
        var menu = new ContextMenuStrip();
        menu.Items.Add(SetMenuGroup("&ログ"));
        menu.Items.Add(SetMenuItem("&集計", Sum_Click));
        menu.Items.Add(SetMenuItem("&終了", Close_Click));
        
        // タスクトレイ
        var icon = new NotifyIcon();
        icon.Icon = SystemIcons.Application;
        icon.Visible = true;
        icon.Text = this.Text;
        icon.ContextMenuStrip = menu;
        
        // バルーンチップ
        icon.BalloonTipTitle = this.Text;
        icon.BalloonTipText = "アクティブウィンドウのプロセスを監視します。";
        icon.ShowBalloonTip(3000);
    }
    
    private ToolStripMenuItem SetMenuGroup(string text)
    {
        var menuGroup = new ToolStripMenuItem();
        menuGroup.Text = text;
        var name = GetLogName();
        var files = Directory.EnumerateFiles(Environment.CurrentDirectory, "*.log", System.IO.SearchOption.TopDirectoryOnly).OrderByDescending(f => f);
        var cnt = 0;
        foreach(var file in files)
        {
            menuGroup.DropDownItems.Add(SetMenuItem(Path.GetFileName(file), Log_Click));
            cnt++;
            if(cnt >= history)
            {
                break;
            }
        }
        return menuGroup;
    }
    
    private ToolStripMenuItem SetMenuItem(string text, EventHandler handler)
    {
        var menuItem = new ToolStripMenuItem();
        menuItem.Text = text;
        menuItem.Click += handler;
        return menuItem;
    }
    
    private void Log_Click(object sender, EventArgs e)
    {
        var name1 = ((ToolStripMenuItem)sender).Text;
        var name2 = "watch.txt";
        File.Copy(name1, name2, true);
        Process.Start(name2);
    }
    
    private void Sum_Click(object sender, EventArgs e)
    {
        if(Application.OpenForms.Count > 0)
        {
            Application.OpenForms[0].Activate();
        }
        else
        {
            SubForm subForm = new SubForm(encoding);
            subForm.ShowDialog();
        }
    }
    
    private void Close_Click(object sender, EventArgs e)
    {
        if(MessageBox.Show("終了しますか？", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        {
            Application.Exit();
        }
    }
    
    private string GetLogName()
    {
        var name = DateTime.Now.ToString("yyyyMMdd") + ".log";
        if(!File.Exists(name))
        {
            var log = "\"日時\",\"インターバル\",\"ID\",\"プロセス\",\"タイトル\"" + Environment.NewLine;
            File.AppendAllText(name, log, Encoding.GetEncoding(encoding));
        }
        return name;
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
        
        var interval = timer.Interval;
        foreach(var item in config.Elements("interval"))
        {
            Int32.TryParse(item.Value, out interval);
            break;
        }
        timer.Interval = interval;
        
        foreach(var item in config.Elements("encoding"))
        {
            encoding = item.Value;
            break;
        }
        
        var num = history;
        foreach(var item in config.Elements("history"))
        {
            Int32.TryParse(item.Value, out num);
            break;
        }
        history = num;
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
                    new XElement("encoding", "shift_jis"),
                    new XElement("history", 10)
                )
            );
            xml.Save(file);
            return xml;
        }
    }
}

class SubForm : Form
{
    DateTimePicker dtp1 = new DateTimePicker();
    DateTimePicker dtp2 = new DateTimePicker();
    
    private string encoding;
    
    public SubForm(string encoding)
    {
        // フォーム
        this.Text = "WATCH";
        this.Size = new Size(260, 100);
        this.MaximizeBox = false;
        this.MinimizeBox = true;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        
        // 日付入力①
        dtp1.Location = new Point(10, 10);
        dtp1.Size = new Size(105, 20);
        dtp1.Value = DateTime.Today.AddDays(-6);
        dtp1.CloseUp += Dtp1_CloseUp;
        this.Controls.Add(dtp1);
        
        // ～
        Label label1 = new Label();
        label1.Location = new Point(120, 15);
        label1.Size = new Size(20, 20);
        label1.Text = "～";
        this.Controls.Add(label1);
        
        // 日付入力②
        dtp2.Location = new Point(140, 10);
        dtp2.Size = new Size(105, 20);
        dtp2.Value = DateTime.Today;
        dtp2.CloseUp += Dtp2_CloseUp;
        this.Controls.Add(dtp2);
        
        // 集計ボタン
        Button button1 = new Button();
        button1.Location = new Point(10, 40);
        button1.Size = new Size(235, 20);
        button1.Text = "集計";
        button1.Click += Button_Click;
        this.Controls.Add(button1);
        
        this.encoding = encoding;
    }
    
    private void Dtp1_CloseUp(object sender, EventArgs e)
    {
        if(dtp1.Value > dtp2.Value)
        {
            dtp2.Value = dtp1.Value;
        }
    }
    
    private void Dtp2_CloseUp(object sender, EventArgs e)
    {
        if(dtp1.Value > dtp2.Value)
        {
            dtp1.Value = dtp2.Value;
        }
    }
    
    private void Button_Click(object sender, EventArgs e)
    {
        var d1 = dtp1.Value.ToString("yyyyMMdd") + ".log";
        var d2 = dtp2.Value.ToString("yyyyMMdd") + ".log";
        
        Dictionary<Tuple<string, string>, long> table = new Dictionary<Tuple<string, string>, long>();
        
        foreach(var file in Directory.EnumerateFiles(".", "*.log"))
        {
            var name = Path.GetFileName(file);
            if(string.Compare(d1, name) <= 0 && string.Compare(d2, name) >= 0)
            {
                var parser = new TextFieldParser(name, Encoding.GetEncoding(encoding));
                using(parser)
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.TrimWhiteSpace = true;
                    while(!parser.EndOfData)
                    {
                        string[] row = parser.ReadFields();
                        if(row.Count() < 5)
                        {
                            continue;
                        }
                        long interval;
                        if(!long.TryParse(row[1], out interval))
                        {
                            continue;
                        }
                        var key = Tuple.Create(row[3], row[4]);
                        if(table.ContainsKey(key))
                        {
                            table[key] += interval;
                        }
                        else
                        {
                            table.Add(key, interval);
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
        foreach(var row in table.OrderBy(r => r.Key.Item1).ThenBy(r => r.Key.Item2))
        {
            var hour = Math.Ceiling(row.Value / 3600000.0 * 100.0) / 100.0;
            var log = "\"" + row.Key.Item1 + "\",\"" + row.Key.Item2 + "\",\"" + hour + "\"" + Environment.NewLine;
            File.AppendAllText(name, log, Encoding.GetEncoding(encoding));
        }
        Process.Start(name);
    }
}
