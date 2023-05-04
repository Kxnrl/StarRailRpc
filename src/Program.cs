using DiscordRPC;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarRailDiscordRpc;

internal static class Program
{
    private const string AppId_Zh = "1100788944901247026";
    private const string AppId_En = "1100789242029948999";

    [STAThread]
    static void Main()
    {
        using var self = new Mutex(true, "StarRail DiscordRPC", out var allow);
        if (!allow)
        {
            MessageBox.Show("StarRail DiscordRPC is already running.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(-1);
        }

        if (Properties.Settings.Default.IsFirstTime)
        {
            AutoStart.Set();
            Properties.Settings.Default.IsFirstTime = false;
            Properties.Settings.Default.Save();
        }

        Task.Run(async () =>
        {
            using var clientZh = new DiscordRpcClient(AppId_Zh);
            using var clientEn = new DiscordRpcClient(AppId_En);
            clientZh.Initialize();
            clientEn.Initialize();

            var playing = false;

            while (true)
            {
                await Task.Delay(1000);

                Debug.Print($"InLoop");

                var handleZh = FindWindow("UnityWndClass", "崩坏：星穹铁道"); // "Honkai: Star Rail"
                var handleEn = FindWindow("UnityWndClass", "Honkai: Star Rail");
                var handle = IntPtr.Zero;

                if (handleZh == IntPtr.Zero)
                {
                    handle = handleEn;
                } else if (handleEn == IntPtr.Zero)
                {
                    handle = handleZh;
                }
                
                if (handle == IntPtr.Zero)
                {
                    Debug.Print($"Not found game process.");
                    playing = false;
                    if (clientEn.CurrentPresence != null)
                    {
                        clientEn.ClearPresence();
                    }
                    if (clientZh.CurrentPresence != null)
                    {
                        clientZh.ClearPresence();
                    }
                    continue;
                }

                try
                {
                    var process = Process.GetProcesses().First(x => x.MainWindowHandle == handle);

                    var isGlobal = CheckGlobal(process);

                    Debug.Print($"Check process with {handle} | {process.ProcessName}");

                    if (!isGlobal)
                    {
                        if (!playing)
                        {
                            playing = true;
                            clientZh.UpdateRpc("logo", "崩坏：星穹铁道");
                            Debug.Print($"Set RichPresence to {process.ProcessName}");
                        }
                        else
                        {
                            Debug.Print($"Keep RichPresence to {process.ProcessName}");
                        }
                    }
                    else
                    {
                        if (!playing)
                        {
                            playing = true;
                            clientEn.UpdateRpc("logo", "Honkai: Star Rail");
                            Debug.Print($"Set RichPresence to  {process.ProcessName}");
                        }
                        else
                        {
                            Debug.Print($"Keep RichPresence to {process.ProcessName}");
                        }
                    }
                }
                catch (Exception e)
                {
                    playing = false;
                    if (clientEn.CurrentPresence != null)
                    {
                        clientEn.ClearPresence();
                    }
                    if (clientZh.CurrentPresence != null)
                    {
                        clientZh.ClearPresence();
                    }
                    Debug.Print($"{e.Message}{Environment.NewLine}{e.StackTrace}");
                }

                GC.Collect();
                GC.WaitForFullGCComplete();
            }
        });

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var notifyMenu = new ContextMenu();
        var exitButton = new MenuItem("Exit");
        var autoButton = new MenuItem("AutoStart" + "    " + (AutoStart.Check() ? "√" : "✘"));
        notifyMenu.MenuItems.Add(0, autoButton);
        notifyMenu.MenuItems.Add(1, exitButton);

        var notifyIcon = new NotifyIcon()
        {
            BalloonTipIcon = ToolTipIcon.Info,
            ContextMenu = notifyMenu,
            Text = "StarRail DiscordRPC",
            Icon = Properties.Resources.tray,
            Visible = true,
        };

        exitButton.Click += (sender, args) =>
        {
            notifyIcon.Visible = false;
            Thread.Sleep(100);
            Environment.Exit(0);
        };
        autoButton.Click += (sender, args) =>
        {
            if (AutoStart.Check())
            {
                AutoStart.Remove();
            }
            else
            {
                AutoStart.Set();
            }

            autoButton.Text = "AutoStart" + "    " + (AutoStart.Check() ? "√" : "✘");
        };


        Application.Run();
    }

    private static void UpdateRpc(this DiscordRpcClient client, string key, string text)
        => client.SetPresence(new RichPresence
        {
            Assets = new Assets
            {
                LargeImageKey = key,
                LargeImageText = text,
            },
            Timestamps = Timestamps.Now,
        });

    private static bool CheckGlobal(Process process)
    {
        var path = GetPathOfWindow(process);
        var file = Path.Combine(Path.GetDirectoryName(path)!, "config.ini");

        if (!File.Exists(file))
            throw new FileNotFoundException("Config not found");

        var value = GetIniSectionValue(file, "General", "cps", "gw_PC");

        return value is not ("gw_PC" or "bilibili_PC");
    }

    private static string GetIniSectionValue(string file, string section, string key, string defaultVal = null)
    {
        var stringBuilder = new StringBuilder(1024);
        GetPrivateProfileString(section, key, defaultVal, stringBuilder, 1024, file);
        return stringBuilder.ToString();
    }

    [DllImport("kernel32.dll")]
    private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    private static string GetPathOfWindow(Process process)
    {
        try
        {
            return process.MainModule!.FileName;
        }
        catch (Exception e)
        {
            Debug.Print(e.ToString());
        }

        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT ProcessId, ExecutablePath, CommandLine FROM Win32_Process");
            using var collection = searcher.Get();
            var results = collection.Cast<ManagementObject>();

            foreach (var item in results)
            {
                var id = (int)(uint)item["ProcessId"];
                var path = (string)item["ExecutablePath"];

                if (id == process.Id)
                {
                    return path;
                }
            }
        }
        catch (Exception e)
        {
            Debug.Print(e.ToString());
        }

        throw new InvalidOperationException("Failed to get path of handle");
    }
}
