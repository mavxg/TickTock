using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using TickTock.Properties;
using System.Windows.Automation;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using System.Drawing;
using System.Drawing.Imaging;

namespace TickTock
{
    class Capture : IDisposable
    {
        Timer timer;
        NotifyIcon ni;
        ToolStripMenuItem toggle;
        SessionSwitchEventHandler sessionSwitchEventHandler;
        LASTINPUTINFO lastInputInfo;

        string logPath;
        string activityFile;
        string lastDay;

        bool restart = false;

        string lastUrl = "not-yo-mama's-sentinal";
        string lastTitle = "not-yo-mama's-sentinal";
        string lastProcess = "not-yo-mama's-sentinal";


        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        private static string GetExecutablePathAboveVista(int ProcessId)
        {
            var buffer = new StringBuilder(1024);
            IntPtr hprocess = OpenProcess(0x1000,
                                          false, ProcessId);
            if (hprocess != IntPtr.Zero)
            {
                try
                {
                    int size = buffer.Capacity;
                    if (QueryFullProcessImageName(hprocess, 0, buffer, out size))
                    {
                        return buffer.ToString();
                    }
                }
                finally
                {
                    CloseHandle(hprocess);
                }
            }
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        [DllImport("kernel32.dll")]
        private static extern bool QueryFullProcessImageName(IntPtr hprocess, int dwFlags,
                       StringBuilder lpExeName, out int size);
        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(int dwDesiredAccess,
                       bool bInheritHandle, int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hHandle);

        public static string GetChromeUrl(Process process)
        {
            if (process == null)
                throw new ArgumentNullException("process");

            if (process.MainWindowHandle == IntPtr.Zero)
                return null;

            AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
            if (element == null)
                return null;

            AutomationElement edit = element.FindFirst(TreeScope.Subtree,
                 new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, "address and search bar", PropertyConditionFlags.IgnoreCase),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)));

            return ((ValuePattern)edit.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
        }

        public static AutomationElement GetEdgeCommandsWindow(AutomationElement edgeWindow)
        {
            try
            {
                return edgeWindow.FindFirst(TreeScope.Children, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                    new PropertyCondition(AutomationElement.NameProperty, "Microsoft Edge")));
            } catch
            {
                return null;
            }
        }

        public static string GetEdgeUrl(AutomationElement edgeCommandsWindow)
        {
            var adressEditBox = edgeCommandsWindow.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "addressEditBox"));

            return ((TextPattern)adressEditBox.GetCurrentPattern(TextPattern.Pattern)).DocumentRange.GetText(int.MaxValue);
        }

        static string GetEdgeUrl(Process process)
        {
            if (process == null)
                throw new ArgumentNullException("process");

            if (process.MainWindowHandle == IntPtr.Zero)
                return null;

            AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
            if (element == null)
                return null;

            AutomationElement window = GetEdgeCommandsWindow(element);
            if (window == null) // not edge
                return null;

            return GetEdgeUrl(window);
        }

        public static string GetFirefoxUrl(Process process)
        {
            if (process == null)
                throw new ArgumentNullException("process");

            if (process.MainWindowHandle == IntPtr.Zero)
                return null;

            AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
            if (element == null)
                return null;


            element = element.FindFirst(TreeScope.Subtree,
                  new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, "search or enter address", PropertyConditionFlags.IgnoreCase),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)));


            if (element == null)
                return null;

            return ((ValuePattern)element.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
        }

        public static string GetInternetExplorerUrl(Process process)
        {
            if (process == null)
                throw new ArgumentNullException("process");

            if (process.MainWindowHandle == IntPtr.Zero)
                return null;

            AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
            if (element == null)
                return null;

            AutomationElement rebar = element.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "ReBarWindow32"));
            if (rebar == null)
                return null;

            AutomationElement edit = rebar.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

            return ((ValuePattern)edit.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
        }

        public static string GetInternetUrl()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                Process p = Process.GetProcessById((int)pid);

                if (p == null)
                {
                    return null;
                }

                switch (p.ProcessName)
                {
                    case "iexplore":
                        return GetInternetExplorerUrl(p);
                    case "chrome":
                        return GetChromeUrl(p);
                    case "firefox":
                        return GetFirefoxUrl(p);
                    default:
                        return GetEdgeUrl(p);
                }
            } catch
            {
                return null;
            }
        }


        string GetActiveProcessFileName()
        {
            IntPtr hwnd = GetForegroundWindow();
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            Process p = Process.GetProcessById((int)pid);

            //If running on Vista or later use the new function
            if (Environment.OSVersion.Version.Major >= 6)
            {
                return GetExecutablePathAboveVista((int)pid);
            }
     
            try
            {
                return p.MainModule.FileName;
            } catch
            {
                return p.ProcessName;
            }
        }

        public int Interval
        {
            get
            {
                return timer.Interval;
            }
            set
            {
                timer.Interval = value;
            }
        }

        public bool Enabled
        {
            get
            {
                return timer.Enabled;
            }
        }

        ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);
        EncoderParameters myEncoderParameters;

        public object ImageUtils { get; private set; }

        public Capture()
        {
            sessionSwitchEventHandler = new SessionSwitchEventHandler(SwitchHandler);
            SystemEvents.SessionSwitch += sessionSwitchEventHandler;

            myEncoderParameters = new EncoderParameters(1);
            myEncoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 50L);

            lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
            lastInputInfo.dwTime = 0;

            ni = new NotifyIcon();
            ni.Click += Ni_Click;

            logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"TickTock");
            activityFile = Path.Combine(logPath, "activity.js");

            System.IO.Directory.CreateDirectory(logPath); //create the log directory if it doesn't exist

            var indexHtml = Path.Combine(logPath, "index.html");
            if (!File.Exists(indexHtml))
            {
                File.WriteAllText(indexHtml, Resources.index_html);
            }

            if (File.Exists(activityFile)) {
                //scan activity file to get the last lastDay (can scan backwards)
                var lines = File.ReadAllLines(activityFile).Reverse().Take(6000);
                foreach (string line in lines)
                {
                    if (line.StartsWith("b=a['"))
                    {
                        lastDay = line.Substring(5, 10);
                        Console.WriteLine("LastDay: " + lastDay);
                        break;
                    }
                }
            }
            

            timer = new Timer();
            timer.Tick += new EventHandler(TimerEvent);
            timer.Interval = 15000; //15 seconds default

            toggle = new ToolStripMenuItem();
            toggle.Text = "Start";
            toggle.Click += new System.EventHandler(Toggle_Click);
            toggle.Image = Resources.Start;
        }

        private void Ni_Click(object sender, EventArgs e)
        {
            var ee = (MouseEventArgs)e;
            if (ee.Button == MouseButtons.Left)
            {
                var indexHtml = Path.Combine(logPath, "index.html");
                System.Diagnostics.Process.Start(indexHtml);
            }
        }

        private void SwitchHandler(object sender, SessionSwitchEventArgs e)
        {
            switch (e.Reason)
            {
                case SessionSwitchReason.SessionLock:
                case SessionSwitchReason.SessionLogoff:
                    restart = timer.Enabled;
                    if (restart) Stop();
                    Console.WriteLine("Lock Encountered");
                    break;
                case SessionSwitchReason.SessionUnlock:
                case SessionSwitchReason.SessionLogon:
                    if (!timer.Enabled && restart) Start();
                    Console.WriteLine("UnLock Encountered");
                    break;
            }
        }

        private ContextMenuStrip CreateContextMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem item;
            ToolStripSeparator sep;

            item = new ToolStripMenuItem();
            item.Text = "Exit";
            item.Text = "Exit";
            item.Click += new System.EventHandler(Exit_Click);
            item.Image = Resources.Exit;
            menu.Items.Add(item);

            sep = new ToolStripSeparator();
            menu.Items.Add(sep);

            menu.Items.Add(toggle);

            return menu;
        }

        public void Display()
        {
            //ni.MouseClick += new MouseEventHandler(Toggle_Click);
            ni.Icon = Resources.SystemTrayIcon;
            ni.Text = "TickTock";
            ni.Visible = true;

            ni.ContextMenuStrip = CreateContextMenu();
        }

        public void Start()
        {
            timer.Start();
            toggle.Text = "Stop";
            toggle.Image = Resources.Stop;
        }

        public void Stop()
        {
            timer.Stop();
            toggle.Text = "Start";
            toggle.Image = Resources.Start;
        }

        private void Toggle_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Toggle clicked");
            if (Enabled)
            {
                Stop();
            } else
            {
                Start();
            }
        }

        void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        public void Dispose()
        {
            ni.Dispose();
            timer.Dispose();
            SystemEvents.SessionSwitch -= sessionSwitchEventHandler;
        }

        const int SPI_GETSCREENSAVERRUNNING = 114;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SystemParametersInfo(
           int uAction, int uParam, ref bool lpvParam,
           int flags);


        // Returns TRUE if the screen saver is actually running
        public static bool GetScreenSaverRunning()
        {
            bool isRunning = false;

            SystemParametersInfo(SPI_GETSCREENSAVERRUNNING, 0,
               ref isRunning, 0);
            return isRunning;
        }

        [DllImport("user32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        uint GetIdleTime() //seconds idle
        {
            if (!GetLastInputInfo(ref lastInputInfo))
            {
                return 0;
            }
            return ((uint)Environment.TickCount - lastInputInfo.dwTime) / 1000;

        }

        private void TimerEvent(Object obj, EventArgs args)
        {

            if (GetScreenSaverRunning()) return; //Don't log anything with screensaver running

            //don't log anything if we have been idle for more than 3 minutes
            if (GetIdleTime() > 180) return;

            IntPtr hwnd = GetForegroundWindow();
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);

            if (pid == 0)
            {
                Console.WriteLine("Locked");
                return;
            }

            var now = DateTime.Now;
            var day = now.ToString("yyyy-MM-dd");
            var time = now.ToString("HH.mm.ss.ffff");

            var title = GetActiveWindowTitle();
            var process = GetActiveProcessFileName();
            var url = GetInternetUrl();

            var snapPath = Path.Combine(logPath, day);
            System.IO.Directory.CreateDirectory(snapPath);

            int duration = (int)(Interval / 1000);

            using (var tw = new StreamWriter(activityFile, true))
            {
                if (tw.BaseStream.Position == 0)
                {
                    //create file if it doesn't exist
                    tw.WriteLine("a={}");
                    lastDay = null;
                }
                if (day != lastDay)
                {
                    lastDay = day;
                    tw.WriteLine("b=a['" + lastDay + "']={}");
                }

                if (title != lastTitle)
                {
                    lastTitle = title;
                    if (title == null)
                        tw.WriteLine("t=null");
                    else
                        tw.WriteLine("t='" + title.Replace("\\", "\\\\").Replace("'", "\\'") + "'");
                }

                if (process != lastProcess)
                {
                    lastProcess = process;
                    if (process == null)
                        tw.WriteLine("p=null");
                    else
                        tw.WriteLine("p='" + process.Replace("\\", "\\\\").Replace("'", "\\'") + "'");
                }

                if (url != lastUrl)
                {
                    lastUrl = url;
                    if (url == null)
                        tw.WriteLine("u=null");
                    else
                        tw.WriteLine("u='" + url.Replace("\\", "\\\\").Replace("'", "\\'") + "'");
                }

                tw.WriteLine("b['{0}']={{t:t,p:p,u:u,d:{1}}}", time, duration);
                tw.Close();
            }

            using (Bitmap screenshot = new Bitmap(SystemInformation.VirtualScreen.Width,
                               SystemInformation.VirtualScreen.Height,
                               PixelFormat.Format32bppArgb))
            using (Graphics screenGraph = Graphics.FromImage(screenshot))
            {

                screenGraph.CopyFromScreen(SystemInformation.VirtualScreen.X,
                                       SystemInformation.VirtualScreen.Y,
                                       0,
                                       0,
                                       SystemInformation.VirtualScreen.Size,
                                       CopyPixelOperation.SourceCopy);
                
                screenshot.Save(Path.Combine(snapPath, time + ".jpg"), jgpEncoder, myEncoderParameters);
            }
            //TODO
            /*
             * Choose older day
             * 
             * Playback slider
             * 
             * Show explorer.exe file path (as url)
             */
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }
    }
}
