using DiscordRPC;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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

                var hndlZh = FindWindow("UnityWndClass", "崩坏：星穹铁道");
                var hndlEn = FindWindow("UnityWndClass", "Honkai: Star Rail");
                if (hndlZh == IntPtr.Zero &&
                    hndlEn == IntPtr.Zero)
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
                    var procEn = Process.GetProcesses().FirstOrDefault(x => x.MainWindowHandle == hndlEn);
                    var procZh = Process.GetProcesses().FirstOrDefault(x => x.MainWindowHandle == hndlZh);

                    Debug.Print($"Check process with {hndlZh} | {procZh?.ProcessName} || {hndlEn} | {procEn?.ProcessName}");

                    if (procZh != null)
                    {
                        if (!playing)
                        {
                            playing = true;

                            clientZh.SetPresence(new RichPresence
                            {
                                Assets = new Assets
                                {
                                    LargeImageKey = "logo",
                                    LargeImageText = "崩坏：星穹铁道",
                                },
                                Timestamps = Timestamps.Now,
                            });

                            Debug.Print($"Set RichPresence to {procZh?.ProcessName}");
                        }
                        else
                        {
                            Debug.Print($"Keep RichPresence to {procZh?.ProcessName}");
                        }
                    }
                    else if (procEn != null)
                    {
                        if (!playing)
                        {
                            playing = true;

                            clientEn.SetPresence(new RichPresence
                            {
                                Assets = new Assets
                                {
                                    LargeImageKey = "logo",
                                    LargeImageText = "Honkai: Star Rail",
                                },
                                Timestamps = Timestamps.Now,
                            });

                            Debug.Print($"Set RichPresence to  {procEn?.ProcessName}");
                        }
                        else
                        {
                            Debug.Print($"Keep RichPresence to {procEn?.ProcessName}");
                        }
                    }
                    else
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
                        Debug.Print($"Clear RichPresence");
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

    [DllImport("user32.dll", EntryPoint = "FindWindow")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
