using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using TickTock.Properties;
using System.Windows.Automation;

namespace TickTock
{
    class Capture : IDisposable
    {
        Timer timer;
        NotifyIcon ni;
        ToolStripMenuItem toggle;

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
            IntPtr hwnd = GetForegroundWindow();
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            Process p = Process.GetProcessById((int)pid);

            if (p == null)
            {
                return null;
            }
            
            switch (p.ProcessName) {
                case "iexplore":
                    return GetInternetExplorerUrl(p);
                case "chrome":
                    return GetChromeUrl(p);
                case "firefox":
                    return GetFirefoxUrl(p);
                default:
                    return GetEdgeUrl(p);
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

        public Capture()
        {
            ni = new NotifyIcon();

            timer = new Timer();
            timer.Tick += new EventHandler(TimerEvent);
            timer.Interval = 15000; //15 seconds default

            toggle = new ToolStripMenuItem();
            toggle.Text = "Start";
            toggle.Click += new System.EventHandler(Toggle_Click);
            toggle.Image = Resources.Start;
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
            Console.WriteLine("Start clicked");
            timer.Start();
            toggle.Text = "Stop";
            toggle.Image = Resources.Stop;
        }

        public void Stop()
        {
            Console.WriteLine("Stop clicked");
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
        }

        private void TimerEvent(Object obj, EventArgs args)
        {
            IntPtr hwnd = GetForegroundWindow();
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);

            if (pid == 0)
            {
                Console.WriteLine("Locked");
                return;
            }

            Console.WriteLine(GetActiveWindowTitle());
            Console.WriteLine(GetActiveProcessFileName());
            Console.WriteLine(GetInternetUrl());
            Console.WriteLine("\n\n");

            //TODO
            /*
             * Write out to AppData\Local\TickTock\YYYY-MM-DD\log.js
             * activity = {}
             * activity['HH.MM.SS.milliseconds'] = { .... stuff we know about this ... }
             * 
             * ??? Can we then also put an index.html object into that folder ???
             * and then do some form of slider/playback for the data?
             */
        }
    }
}
