






using System;
using System.Collections;
using System.Configuration.Install;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Principal;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Forms;

namespace ThingsManager
{
    [RunInstaller(true)]
    public class ThmServiceInstaller : Installer
    {
        public ThmServiceInstaller()
        {
            ServiceProcessInstaller pi = new ServiceProcessInstaller();
            pi.Account = ServiceAccount.LocalSystem;
            ServiceInstaller si = new ServiceInstaller();
            si.ServiceName = "ThingsManager";
            si.DisplayName = "ThingsManager \u00b7 轻量仓库管理";
            si.Description = "ThingsManager 轻量仓库管理后台服务（监听 0.0.0.0:3200 供局域网访问；托盘/设置页可控制开机自启）。";
            si.StartType = ServiceStartMode.Automatic;
            Installers.Add(pi);
            Installers.Add(si);
        }
    }

    internal static class Paths
    {
        public static string ExeDir
        {
            get { return AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\') + "\\"; }
        }
        public static string AppDir { get { return Path.Combine(ExeDir, "app"); } }
        public static string NodeExe { get { return Path.Combine(ExeDir, "runtime", "node.exe"); } }
        public static string Supervisor { get { return Path.Combine(AppDir, "supervisor.js"); } }
        public static string CfgFile { get { return Path.Combine(AppDir, "runtime.config.json"); } }

        
        
        public static string PortableFlag { get { return Path.Combine(ExeDir, "portable.flag"); } }
        public static bool Portable { get { try { return File.Exists(PortableFlag); } catch { return false; } } }

        public static int ReadPort()
        {
            int port = 0;
            try
            {
                if (File.Exists(CfgFile))
                {
                    string text = File.ReadAllText(CfgFile);
                    int key = text.IndexOf("\"port\"", StringComparison.OrdinalIgnoreCase);
                    if (key >= 0)
                    {
                        int colon = text.IndexOf(':', key);
                        if (colon >= 0)
                        {
                            int start = colon + 1;
                            while (start < text.Length && !char.IsDigit(text[start])) start++;
                            int end = start;
                            while (end < text.Length && char.IsDigit(text[end])) end++;
                            if (int.TryParse(text.Substring(start, end - start), out port) && port > 0)
                            {
                                return port;
                            }
                        }
                    }
                }
            }
            catch { }
            string env = Environment.GetEnvironmentVariable("THM_PORT");
            if (!string.IsNullOrEmpty(env) && int.TryParse(env, out port) && port > 0)
            {
                return port;
            }
            return 3200;
        }
    }

    internal class ThmService : ServiceBase
    {
        private Process _node;

        protected override void OnStart(string[] args)
        {
            try
            {
                if (!File.Exists(Paths.NodeExe)) throw new Exception("未找到内嵌 Node 运行时：" + Paths.NodeExe);
                if (!File.Exists(Paths.Supervisor)) throw new Exception("未找到应用入口：" + Paths.Supervisor);
                ProcessStartInfo psi = new ProcessStartInfo(Paths.NodeExe, "\"" + Paths.Supervisor + "\"");
                psi.WorkingDirectory = Paths.AppDir;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                _node = Process.Start(psi);
                if (_node == null) throw new Exception("无法启动守护进程");
            }
            catch (Exception ex)
            {
                throw new Exception("ThingsManager 服务启动失败：" + ex.Message, ex);
            }
        }

        protected override void OnStop()
        {
            try
            {
                if (_node != null && !_node.HasExited)
                {
                    ThmService.KillTree(_node.Id);
                }
            }
            catch { }
        }

        protected override void OnShutdown() { OnStop(); }

        internal static void KillTree(int pid)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("taskkill", "/PID " + pid + " /T /F");
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                using (Process p = Process.Start(psi))
                {
                    if (p != null) p.WaitForExit(3000);
                }
            }
            catch { }
        }
    }

    internal static class Ops
    {
        public static bool IsAdmin()
        {
            try
            {
                return new WindowsPrincipal(WindowsIdentity.GetCurrent())
                    .IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch { return false; }
        }

        public static bool ServiceExists(string name)
        {
            try
            {
                foreach (ServiceController s in ServiceController.GetServices())
                {
                    if (s.ServiceName == name) return true;
                }
                return false;
            }
            catch { return false; }
        }

        public static string Sc(string args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("sc", args);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                using (Process p = Process.Start(psi))
                {
                    if (p == null) return "";
                    p.WaitForExit(6000);
                    return p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
                }
            }
            catch (Exception ex) { return "ERR " + ex.Message; }
        }

        public static bool ServiceStartTypeAuto()
        {
            string o = Ops.Sc("qc ThingsManager").ToLowerInvariant();
            bool hasStartType = o.Contains("start_type");
            if (hasStartType && (o.Contains("4   auto_start") || o.Contains("auto_start") || o.Contains("2   auto_start"))) return true;
            if (hasStartType && (o.Contains("demand_start"))) return false;
            if (o.Contains("disabled")) return false;
            return true;
        }

        public static bool RunElevated(string args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo(Application.ExecutablePath, args);
                psi.Verb = "runas";
                psi.UseShellExecute = true;
                Process.Start(psi);
                return true;
            }
            catch { return false; }
        }

        public static void Info(string text)
        {
            MessageBox.Show(text, "ThingsManager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    internal static class AppMain
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                string first = args[0].ToLowerInvariant();
                if (first == "--install") return AppMain.InstallService(false);
                if (first == "--install-quiet") return AppMain.InstallService(true);
                if (first == "--uninstall") return AppMain.UninstallService(false);
                if (first == "--uninstall-quiet") return AppMain.UninstallService(true);
                if (first == "--restart") { AppMain.RestartService(); return 0; }
                if (first == "--autostart" && args.Length > 1)
                {
                    AppMain.SetAutostart(args[1]);
                    return 0;
                }
                if (first == "--start-svc") { AppMain.StartServiceOnly(); return 0; }
                if (first == "--stop-svc") { AppMain.StopServiceOnly(); return 0; }
            }

            if (Environment.UserInteractive)
            {
                
                
                bool openMode = true;
                if (args != null && args.Length > 0 && args[0].ToLowerInvariant() == "--tray") openMode = false;
                
                if (Paths.Portable) return AppMain.RunPortable(openMode);
                return AppMain.RunInteractive(openMode);
            }

            ServiceBase.Run(new ThmService());
            return 0;
        }

        
        private static int RunInteractive(bool openMode)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (openMode)
            {
                AppMain.EnsureServiceRunning();   
                ThmTray.OpenPanel();              
            }
            
            bool createdNew;
            using (System.Threading.Mutex trayMutex = new System.Threading.Mutex(true, "ThingsManager_Tray", out createdNew))
            {
                if (!createdNew) return 0; 
                Application.Run(new ThmTray());
            }
            return 0;
        }

        




        private static Process _portableNode;

        private static int RunPortable(bool openMode)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            int port = Paths.ReadPort();
            
            if (!ThmTray.PortUp(port)) AppMain.StartPortableNode();
            if (openMode) ThmTray.OpenPanel();
            bool createdNew;
            using (System.Threading.Mutex trayMutex = new System.Threading.Mutex(true, "ThingsManager_Tray", out createdNew))
            {
                if (!createdNew) return 0;   
                Application.Run(new ThmTray(true));
            }
            return 0;
        }

        
        private static readonly object _logLock = new object();
        private static string _logFile;

        private static void AppendLog(string line)
        {
            if (string.IsNullOrEmpty(line)) return;
            try
            {
                if (_logFile == null)
                {
                    _logFile = Path.Combine(Paths.ExeDir, "logs", "server.log");
                    Directory.CreateDirectory(Path.GetDirectoryName(_logFile));
                    FileInfo fi = new FileInfo(_logFile);
                    if (fi.Exists && fi.Length > 2 * 1024 * 1024) fi.Delete();
                }
                lock (_logLock)
                {
                    File.AppendAllText(_logFile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss  ") + line + Environment.NewLine, System.Text.Encoding.UTF8);
                }
            }
            catch { }
        }

        
        internal static bool StartPortableNode()
        {
            try
            {
                if (!File.Exists(Paths.NodeExe))
                {
                    Ops.Info("未找到内嵌运行环境：\n" + Paths.NodeExe + "\n\n请把压缩包完整解压到本地文件夹后再双击运行（不要在压缩包内直接打开）。");
                    return false;
                }
                if (!File.Exists(Paths.Supervisor))
                {
                    Ops.Info("未找到程序入口：\n" + Paths.Supervisor);
                    return false;
                }
                ProcessStartInfo psi = new ProcessStartInfo(Paths.NodeExe, "\"" + Paths.Supervisor + "\"");
                psi.WorkingDirectory = Paths.AppDir;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                
                
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                
                psi.StandardOutputEncoding = System.Text.Encoding.UTF8;
                psi.StandardErrorEncoding = System.Text.Encoding.UTF8;
                _portableNode = Process.Start(psi);
                if (_portableNode == null) return false;
                _portableNode.OutputDataReceived += delegate(object s, DataReceivedEventArgs e) { AppMain.AppendLog(e.Data); };
                _portableNode.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e) { AppMain.AppendLog("[stderr] " + e.Data); };
                _portableNode.BeginOutputReadLine();
                _portableNode.BeginErrorReadLine();
                return true;
            }
            catch (Exception ex)
            {
                Ops.Info("启动失败：" + ex.Message);
                return false;
            }
        }

        
        internal static void StopPortableNode()
        {
            try
            {
                if (_portableNode != null && !_portableNode.HasExited)
                {
                    ProcessStartInfo psi = new ProcessStartInfo("taskkill", "/PID " + _portableNode.Id + " /T /F");
                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;
                    using (Process p = Process.Start(psi)) { if (p != null) p.WaitForExit(8000); }
                }
            }
            catch { }
        }

        
        internal static bool EnsureServiceRunning()
        {
            try
            {
                if (!Ops.ServiceExists("ThingsManager")) return false;
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    sc.Refresh();
                    if (sc.Status == ServiceControllerStatus.Running || sc.Status == ServiceControllerStatus.StartPending) return true;
                    
                    if (sc.Status == ServiceControllerStatus.StopPending)
                    {
                        try { sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(15)); } catch { }
                    }
                    if (!Ops.IsAdmin())
                    {
                        if (!Ops.RunElevated("--start-svc")) return false; 
                    }
                    else
                    {
                        try { sc.Start(); } catch { }
                    }
                    try
                    {
                        sc.Refresh();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(20));
                        return sc.Status == ServiceControllerStatus.Running;
                    }
                    catch { return false; }
                }
            }
            catch { return false; }
        }

        internal static void StartServiceOnly()
        {
            if (!Ops.IsAdmin()) return;
            try
            {
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    sc.Refresh();
                    if (sc.Status != ServiceControllerStatus.Running) sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(25));
                }
            }
            catch { }
        }

        internal static void StopServiceOnly()
        {
            if (!Ops.IsAdmin()) return;
            try
            {
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    sc.Refresh();
                    if (sc.Status != ServiceControllerStatus.Stopped && sc.Status != ServiceControllerStatus.StopPending) sc.Stop();
                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(25));
                }
            }
            catch { }
        }

        
        
        private static bool StopServiceAndWait(int seconds)
        {
            try
            {
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    if (sc.Status == ServiceControllerStatus.Stopped) return true;
                    if (sc.Status != ServiceControllerStatus.StopPending) sc.Stop();
                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(seconds));
                    return sc.Status == ServiceControllerStatus.Stopped;
                }
            }
            catch { return false; }
        }

        
        private static bool StartServiceAndWait(int seconds)
        {
            try
            {
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    if (sc.Status == ServiceControllerStatus.Running) return true;
                    if (sc.Status == ServiceControllerStatus.StartPending)
                    {
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(seconds));
                        return sc.Status == ServiceControllerStatus.Running;
                    }
                    if (sc.Status == ServiceControllerStatus.StopPending)
                    {
                        try { sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(15)); } catch { }
                    }
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(seconds));
                    return sc.Status == ServiceControllerStatus.Running;
                }
            }
            catch { return false; }
        }

        private static int InstallService(bool quiet)
        {
            if (!Ops.IsAdmin())
            {
                if (quiet) return 2;
                Ops.Info("安装 ThingsManager 服务需要管理员权限。");
                if (Ops.RunElevated("--install")) return 0;
                return 1;
            }
            try
            {
                bool existed = Ops.ServiceExists("ThingsManager");
                if (existed)
                {
                    if (!quiet) Ops.Info("检测到 ThingsManager 服务已存在，将停止并重新注册。");
                    
                    StopServiceAndWait(30);
                    
                    
                    try
                    {
                        using (AssemblyInstaller old = new AssemblyInstaller(Application.ExecutablePath, null))
                        {
                            old.UseNewContext = true;
                            old.Uninstall(null);
                        }
                    }
                    catch {  }
                    
                    for (int i = 0; i < 30 && Ops.ServiceExists("ThingsManager"); i++) System.Threading.Thread.Sleep(500);
                }
                using (AssemblyInstaller installer = new AssemblyInstaller(Application.ExecutablePath, null))
                {
                    installer.UseNewContext = true;
                    installer.Install(new Hashtable());
                    installer.Commit(new Hashtable());
                }
                
                bool started = StartServiceAndWait(25);
                if (!started)
                {
                    System.Threading.Thread.Sleep(2000);
                    started = StartServiceAndWait(25);
                }
                if (!quiet)
                {
                    Ops.Info(started
                        ? "ThingsManager 服务安装成功并已启动（监听 0.0.0.0:3200）。\n提示：若端口被占用，请在系统设置里换端口或释放占用后重试。"
                        : "服务已注册，但本次未能自动启动（可能端口被占用或服务正在收尾）。\n可稍后在托盘图标菜单点「重启服务」，或重启计算机。");
                }
                return started ? 0 : 1;
            }
            catch (Exception ex)
            {
                if (!quiet) Ops.Info("安装失败：" + ex.Message);
                return 1;
            }
        }

        private static int UninstallService(bool quiet)
        {
            if (!Ops.IsAdmin())
            {
                if (quiet) return 2;
                Ops.Info("卸载 ThingsManager 服务需要管理员权限。");
                if (Ops.RunElevated("--uninstall")) return 0;
                return 1;
            }
            try
            {
                Ops.Sc("stop ThingsManager");
                using (AssemblyInstaller installer = new AssemblyInstaller(Application.ExecutablePath, null))
                {
                    installer.UseNewContext = true;
                    installer.Uninstall(null);
                }
                if (!quiet) Ops.Info("ThingsManager 服务已卸载。");
                return 0;
            }
            catch (Exception ex)
            {
                if (!quiet) Ops.Info("卸载失败：" + ex.Message);
                return 1;
            }
        }

        internal static void RestartService()
        {
            try
            {
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    if (sc.Status != ServiceControllerStatus.Stopped)
                    {
                        sc.Stop();
                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(20));
                    }
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(25));
                }
            }
            catch (Exception ex)
            {
                Ops.Info("重启失败（请以管理员运行，或确认服务存在且端口未被占用）：" + ex.Message);
            }
        }

        private static void SetAutostart(string mode)
        {
            if (!Ops.IsAdmin())
            {
                Ops.Info("修改开机自启需要管理员权限。");
                return;
            }
            string value = (mode != null && mode.ToLowerInvariant() == "auto") ? "auto" : "demand";
            Ops.Sc("config ThingsManager start= " + value);
        }
    }

    internal class ThmTray : ApplicationContext
    {
        private NotifyIcon _icon;
        private ContextMenuStrip _menu;
        private ToolStripMenuItem _miAutostart;
        private readonly bool _portable;

        public ThmTray() : this(false) { }

        public ThmTray(bool portable)
        {
            _portable = portable;
            _icon = new NotifyIcon();
            _icon.Icon = ThmTray.LoadTrayIcon();
            _icon.Visible = true;

            _menu = new ContextMenuStrip();
            _menu.Items.Add("打开面板", null, OnOpenClick);
            if (_portable)
            {
                
                _menu.Items.Add("打开数据目录", null, OnOpenDataClick);
            }
            else
            {
                _miAutostart = new ToolStripMenuItem("开机自启（服务）", null, OnAutostartClick);
                _menu.Items.Add("重启服务", null, OnRestartClick);
                _menu.Items.Add(_miAutostart);
            }
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_portable ? "关闭程序（结束后台并退出）" : "关闭程序（停止服务并退出）", null, OnExitClick);
            _icon.ContextMenuStrip = _menu;
            _icon.DoubleClick += OnOpenClick;
            
            _icon.MouseUp += delegate(object s, MouseEventArgs e) { if (e.Button == MouseButtons.Left) _menu.Show(Cursor.Position); };
            RefreshState();
            
            _icon.BalloonTipTitle = _portable ? "ThingsManager · 便携版" : "ThingsManager · 托盘守护";
            _icon.BalloonTipText = _portable
                ? "已在本机运行（打开面板 / 打开数据目录 / 关闭程序）。\n数据保存在程序目录的 data 文件夹，拷走整个文件夹即可搬机。\n若看不到图标，请点击任务栏“^”展开隐藏图标。"
                : "已在系统托盘运行（打开面板 / 重启 / 开机自启 / 关闭程序）。\n若看不到图标，请点击任务栏“^”展开隐藏图标。";
            _icon.ShowBalloonTip(2000);
        }

        
        private void OnOpenDataClick(object sender, EventArgs e)
        {
            try
            {
                string dir = Path.Combine(Paths.ExeDir, "data");
                Directory.CreateDirectory(dir);
                Process.Start("explorer.exe", "\"" + dir + "\"");
            }
            catch { }
        }

        
        private static Icon LoadTrayIcon()
        {
            try
            {
                string ico = Path.Combine(Paths.ExeDir, "logo.ico");
                if (File.Exists(ico)) return new Icon(ico);
            }
            catch { }
            return SystemIcons.Application;
        }

        
        public static void OpenPanel()
        {
            try
            {
                int port = Paths.ReadPort();
                string host = ThmTray.PickReadyHost(port, 15000);
                Process.Start("http://" + host + ":" + port + "/");
            }
            catch { }
        }

        
        public static string GetLanIp()
        {
            try
            {
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != OperationalStatus.Up) continue;
                    if (ni.NetworkInterfaceType == NetworkInterfaceType.Loopback) continue;
                    IPInterfaceProperties props = ni.GetIPProperties();
                    if (props.GatewayAddresses == null || props.GatewayAddresses.Count == 0) continue;
                    foreach (UnicastIPAddressInformation u in props.UnicastAddresses)
                    {
                        if (u.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            byte[] b = u.Address.GetAddressBytes();
                            if (b[0] == 127 || b[0] == 169) continue;
                            return u.Address.ToString();
                        }
                    }
                }
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != OperationalStatus.Up) continue;
                    foreach (UnicastIPAddressInformation u in ni.GetIPProperties().UnicastAddresses)
                    {
                        if (u.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            byte[] b = u.Address.GetAddressBytes();
                            if (b[0] == 127 || b[0] == 169) continue;
                            return u.Address.ToString();
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        
        public static string PickReadyHost(int port, int timeoutMs)
        {
            string lan = ThmTray.GetLanIp();
            string[] hosts = (string.IsNullOrEmpty(lan) || lan == "127.0.0.1")
                ? new string[] { "127.0.0.1" }
                : new string[] { lan, "127.0.0.1" };
            DateTime end = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (DateTime.UtcNow < end)
            {
                foreach (string h in hosts) if (ThmTray.CanConnect(h, port, 600)) return h;
                Thread.Sleep(250);
            }
            foreach (string h in hosts) if (ThmTray.CanConnect(h, port, 800)) return h;
            return hosts[0];
        }

        private static bool CanConnect(string host, int port, int timeoutMs)
        {
            try
            {
                using (TcpClient c = new TcpClient())
                {
                    IAsyncResult ar = c.BeginConnect(host, port, null, null);
                    if (ar.AsyncWaitHandle.WaitOne(timeoutMs)) { c.EndConnect(ar); return true; }
                }
            }
            catch { }
            return false;
        }

        
        public static bool PortUp(int port)
        {
            return ThmTray.CanConnect("127.0.0.1", port, 500);
        }

        private bool IsRunning()
        {
            try
            {
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    return sc.Status == ServiceControllerStatus.Running;
                }
            }
            catch { return false; }
        }

        private void RefreshState()
        {
            try
            {
                if (_portable)
                {
                    int p = Paths.ReadPort();
                    _icon.Text = "ThingsManager（便携版） · " + (ThmTray.PortUp(p) ? "运行中" : "未运行") + " · :" + p;
                    return;
                }
                bool svc = Ops.ServiceExists("ThingsManager");
                _miAutostart.Enabled = svc;
                _miAutostart.Checked = svc && Ops.ServiceStartTypeAuto();
                string status = svc ? (IsRunning() ? "运行中" : "未运行") : "服务未安装";
                _icon.Text = "ThingsManager \u00b7 " + status + " \u00b7 :" + Paths.ReadPort();
            }
            catch { }
        }

        private void OnOpenClick(object sender, EventArgs e)
        {
            
            if (_portable)
            {
                int p = Paths.ReadPort();
                if (!ThmTray.PortUp(p)) AppMain.StartPortableNode();
            }
            else
            {
                AppMain.EnsureServiceRunning();
            }
            ThmTray.OpenPanel();
            RefreshState();
        }

        private void OnRestartClick(object sender, EventArgs e)
        {
            if (!Ops.IsAdmin())
            {
                if (!Ops.RunElevated("--restart"))
                {
                    Ops.Info("需要管理员权限重启服务。");
                }
                return;
            }
            AppMain.RestartService();
            RefreshState();
        }

        private void OnAutostartClick(object sender, EventArgs e)
        {
            bool want = !_miAutostart.Checked;
            string mode = want ? "auto" : "demand";
            if (!Ops.IsAdmin())
            {
                Ops.RunElevated("--autostart " + mode);
                return;
            }
            Ops.Sc("config ThingsManager start= " + mode);
            RefreshState();
        }

        private void OnExitClick(object sender, EventArgs e)
        {
            
            string msg = _portable
                ? "确定要关闭 ThingsManager 吗？\n\n将结束后台进程并退出托盘（网页将无法访问）。数据仍保存在程序目录的 data 文件夹里。"
                : "确定要关闭 ThingsManager 吗？\n\n将一并停止后台服务并退出托盘（网页将无法访问）。";
            if (MessageBox.Show(msg, "ThingsManager", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
            if (_portable)
            {
                AppMain.StopPortableNode();
            }
            else if (Ops.IsAdmin())
            {
                AppMain.StopServiceOnly();
            }
            else if (!Ops.RunElevated("--stop-svc"))
            {
                Ops.Info("未取得管理员权限，无法停止 ThingsManager 后台服务。\n\n托盘将退出，后台服务继续运行（网页仍可访问）。");
            }
            _icon.Visible = false;
            _icon.Dispose();
            Application.Exit();
        }
    }
}
