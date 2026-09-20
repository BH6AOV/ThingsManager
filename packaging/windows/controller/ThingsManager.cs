// ThingsManager 桌面控制器（Windows 服务托管 + 系统托盘守护）
// 编译目标：.NET Framework 4.8（Windows 自带，零额外运行时）
// 用法：--install / --uninstall / --restart / --autostart <auto|demand>
//       --open（桌面/开始菜单图标：确保服务运行后自动打开 http://<本机IP>:<端口>）
//       --start-svc / --stop-svc（提权启动/停止服务，供托盘调用）
//       --tray（登录自启/安装后拉起：仅托盘，不弹浏览器）
//       无参数且由 SCM 启动 → 作为 Windows 服务运行
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
                // 交互启动：--tray = 仅托盘（登录自启 / 安装后拉起用，不自动弹浏览器）；
                // 桌面 / 开始菜单快捷方式或直接双击（默认，等价 --open）= 确保后台服务在运行，并自动打开 http://<本机IP>:<端口>
                bool openMode = true;
                if (args != null && args.Length > 0 && args[0].ToLowerInvariant() == "--tray") openMode = false;
                return AppMain.RunInteractive(openMode);
            }

            ServiceBase.Run(new ThmService());
            return 0;
        }

        // 交互入口：openMode=true 时先“确保服务运行 + 自动打开面板”，再成为托盘（若已有托盘则仅开面板即退出）
        private static int RunInteractive(bool openMode)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (openMode)
            {
                AppMain.EnsureServiceRunning();   // 修复：托盘“关闭程序”停掉服务后，再启动(桌面图标)时服务未拉起导致端口监听失效
                ThmTray.OpenPanel();              // 自动打开 http://<本机IP>:<端口>
            }
            // 托盘单实例：安装后即时启动 + 登录自启/双击重复拉起时，后到者直接退出，避免图标重复
            bool createdNew;
            using (System.Threading.Mutex trayMutex = new System.Threading.Mutex(true, "ThingsManager_Tray", out createdNew))
            {
                if (!createdNew) return 0; // 同会话已有托盘实例在运行（本次仅打开了面板）
                Application.Run(new ThmTray());
            }
            return 0;
        }

        // 确保 ThingsManager 后台服务处于运行状态；服务未启动且非管理员时自动提权拉起
        internal static bool EnsureServiceRunning()
        {
            try
            {
                if (!Ops.ServiceExists("ThingsManager")) return false;
                using (ServiceController sc = new ServiceController("ThingsManager"))
                {
                    sc.Refresh();
                    if (sc.Status == ServiceControllerStatus.Running || sc.Status == ServiceControllerStatus.StartPending) return true;
                    // “关闭程序”停止可能尚未完全结束：先等它彻底停下，避免 Start 冲突
                    if (sc.Status == ServiceControllerStatus.StopPending)
                    {
                        try { sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(15)); } catch { }
                    }
                    if (!Ops.IsAdmin())
                    {
                        if (!Ops.RunElevated("--start-svc")) return false; // 用户取消提权
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
                if (Ops.ServiceExists("ThingsManager"))
                {
                    if (!quiet) Ops.Info("检测到 ThingsManager 服务已存在，将停止并删除后重新安装。");
                    Ops.Sc("stop ThingsManager");
                }
                using (AssemblyInstaller installer = new AssemblyInstaller(Application.ExecutablePath, null))
                {
                    installer.UseNewContext = true;
                    installer.Install(new Hashtable());
                    installer.Commit(new Hashtable());
                }
                Ops.Sc("start ThingsManager");
                if (!quiet)
                {
                    Ops.Info("ThingsManager 服务安装成功并已启动（监听 0.0.0.0:3200）。\n提示：若端口被占用或启动失败，请以管理员释放该端口后重试。");
                }
                return 0;
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

        public ThmTray()
        {
            _icon = new NotifyIcon();
            _icon.Icon = ThmTray.LoadTrayIcon();
            _icon.Visible = true;

            _menu = new ContextMenuStrip();
            _miAutostart = new ToolStripMenuItem("开机自启（服务）", null, OnAutostartClick);
            _menu.Items.Add("打开面板", null, OnOpenClick);
            _menu.Items.Add("重启服务", null, OnRestartClick);
            _menu.Items.Add(_miAutostart);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add("关闭程序（停止服务并退出）", null, OnExitClick);
            _icon.ContextMenuStrip = _menu;
            _icon.DoubleClick += OnOpenClick;
            // 单击托盘小按钮即弹出菜单（用户反馈原需右键不便）
            _icon.MouseUp += delegate(object s, MouseEventArgs e) { if (e.Button == MouseButtons.Left) _menu.Show(Cursor.Position); };
            RefreshState();
            // 首次出现提示：图标默认在托盘区；若被折叠请点任务栏“^”展开
            _icon.BalloonTipTitle = "ThingsManager · 托盘守护";
            _icon.BalloonTipText = "已在系统托盘运行（打开面板 / 重启 / 开机自启 / 关闭程序）。\n若看不到图标，请点击任务栏“^”展开隐藏图标。";
            _icon.ShowBalloonTip(2000);
        }

        // 托盘图标：优先使用与程序同目录的 logo.ico（默认全局 Logo），缺失时退回系统图标
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

        // 自动打开管理面板：优先 <本机IP>:<端口>（服务就绪才打开），本机 IP 不可达时自动回退 127.0.0.1
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

        // 取本机 IPv4 局域网地址：偏好“已连接且带默认网关”的网卡；无网关时回退任一非回环 IPv4；找不到返回 null（调用方回退 127.0.0.1）
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

        // 轮询端口直到就绪，返回实际可用的主机（本机 IP 优先、127.0.0.1 兜底）
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
            // 先确保服务在运行（未启动则提权拉起），再打开面板
            AppMain.EnsureServiceRunning();
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
            // “关闭程序”：先确认，再一并停止 ThingsManager 后台服务（需管理员时自动提权），避免“退托盘后程序仍在后台运行”
            if (MessageBox.Show("确定要关闭 ThingsManager 吗？\n\n将一并停止后台服务并退出托盘（网页将无法访问）。",
                "ThingsManager", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
            if (Ops.IsAdmin())
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
