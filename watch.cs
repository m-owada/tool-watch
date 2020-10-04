using System;
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
        try
        {
            var mutex = new Mutex(false, "WATCH_COPYRIGHT_2020_M-OWADA");
            if(!mutex.WaitOne(0, false))
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
        var files = Directory.EnumerateFiles(Environment.CurrentDirectory, "*.log", SearchOption.TopDirectoryOnly).OrderByDescending(f => f);
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
            var log = "\"日時\",\"ID\",\"プロセス\",\"タイトル\"" + Environment.NewLine;
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
