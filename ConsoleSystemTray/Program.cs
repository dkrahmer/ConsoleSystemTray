using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConsoleSystemTray
{
	class Program
	{
		private static readonly Lazy<string> _helpOutput = new(() => string.Concat("Usage: ConsoleSystemTray.exe [-p filePath]", "\n",
			"\n",
			"Arguments:", "\n",
			"  -p\tthe console application to start", "\n",
			"\n",
			"Options:", "\n",
			"  -a\tsets the application arguments", "\n",
			"  -d\tsets the application working directory", "\n",
			"  -i\tsets the tray icon", "\n",
			"  -m\tstart minimized", "\n",
			"  -t\tsets the tray icon tooltip text", "\n",
			"  -s\tprevent Windows from entering sleep mode", "\n",
			"  -h\tshow the help message and exit", "\n")
		);

		static void Main(string[] args)
		{
			if (args.Length > 0)
			{
				try
				{
					var cmds = Utils.GetCommondLines(args);
					if (cmds.ContainsKey("-h"))
					{
						ShowHelpInfo();
						return;
					}

					var options = new Options()
					{
						Path = cmds.GetArgument("-p"),
						Arguments = cmds.GetArgument("-a", true),
						BaseDirectory = cmds.GetArgument("-d", true),
						Icon = cmds.GetArgument("-i", true),
						Tip = cmds.GetArgument("-t", true),
						IsPreventSleep = cmds.ContainsKey("-s"),
						IsStartMinimized = cmds.ContainsKey("-m")
					};

					if (Environment.OSVersion.Version.Major >= 6)
					{
						// Windows 11 requires conhost.exe to run console applications
						options.Arguments = options.Path + (string.IsNullOrWhiteSpace(options.Arguments) ? "" : " " + options.Arguments);
						options.Path = @"C:\Windows\System32\conhost.exe";
					}

					Run(options);
				}
				catch (CmdArgumentException e)
				{
					ShowHelpInfo(e.Message);
					Environment.ExitCode = -1;
				}
				catch (Exception e)
				{
					MessageBox.Show(e.ToString(), "Exception", MessageBoxButtons.OK, MessageBoxIcon.Information);
					Environment.ExitCode = -1;
				}
			}
			else
			{
				ShowHelpInfo();
				Environment.ExitCode = -1;
			}
		}

		private static void ShowHelpInfo(string exceptionMessage = null)
		{
			bool isException = !string.IsNullOrEmpty(exceptionMessage);
			MessageBox.Show((isException ? exceptionMessage + "\n\n\n" : "") + _helpOutput.Value, "Help", MessageBoxButtons.OK, isException ? MessageBoxIcon.Error : MessageBoxIcon.Information);
		}

		private static void Run(Options options)
		{
			Icon trayIcon;
			if (!string.IsNullOrEmpty(options.Icon))
			{
				trayIcon = new Icon(options.Icon);
			}
			else
			{
				trayIcon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
			}

			ProcessStartInfo processStartInfo = new ProcessStartInfo()
			{
				FileName = options.Path,
				ErrorDialog = true,
			};
			if (!string.IsNullOrEmpty(options.Arguments))
			{
				processStartInfo.Arguments = options.Arguments;
			}
			if (!string.IsNullOrEmpty(options.BaseDirectory))
			{
				processStartInfo.WorkingDirectory = options.BaseDirectory;
			}

			var p = Process.Start(processStartInfo);
			ExitAtSame(p);
			SwitchWindow(GetConsoleWindow());

			if (options.IsPreventSleep)
			{
				SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS | EXECUTION_STATE.ES_SYSTEM_REQUIRED);
			}

			string trapText = !string.IsNullOrEmpty(options.Tip) ? options.Tip : GetTrayText(p);
			NotifyIcon tray = new NotifyIcon
			{
				Icon = trayIcon,
				Text = trapText,
				BalloonTipTitle = trapText,
				BalloonTipText = trapText,
				Visible = true,
			};

			tray.MouseDoubleClick += (s, e) =>
			{
				SwitchWindow(GetMainWindowHandle(p));
			};

			if (options.IsStartMinimized)
			{
				Task.Run(() =>
				{
					Thread.Sleep(100); // short delay to allow the console to fully initialize
					SwitchWindow(GetMainWindowHandle(p));
				});
			}

			Application.Run();
		}

		private static IntPtr GetMainWindowHandle(Process process)
		{
			IntPtr mainWindowHandle = process.MainWindowHandle;
			if (mainWindowHandle == IntPtr.Zero)
			{
				mainWindowHandle = GetMainWindowHandleFromProcessId(process.Id);
			}
			return mainWindowHandle;
		}

		private static IntPtr GetMainWindowHandleFromProcessId(int processId)
		{
			ProcessThreadCollection processThreads = Process.GetProcessById(processId).Threads;
			foreach (ProcessThread thread in processThreads)
			{
				IntPtr mainWindowHandle = GetMainWindowHandleFromThreadId(thread.Id);
				if (mainWindowHandle != IntPtr.Zero)
				{
					return mainWindowHandle;
				}
			}
			return IntPtr.Zero;
		}

		private static IntPtr GetMainWindowHandleFromThreadId(int threadId)
		{
			const int GW_OWNER = 4;
			IntPtr mainWindowHandle = IntPtr.Zero;
			EnumThreadWindows(threadId, (hWnd, lParam) =>
			{
				if (GetWindow(hWnd, GW_OWNER) == IntPtr.Zero)
				{
					mainWindowHandle = hWnd;
					return false;
				}
				return true;
			}, IntPtr.Zero);
			return mainWindowHandle;
		}

		private static string GetTrayText(Process p)
		{
			const int kMaxCount = 63;
			Thread.Sleep(500);  // make MainWindowTitle not empty 
			string title = p.MainWindowTitle; ;
			if (title.Length > kMaxCount)
			{
				title = title.Substring(title.Length - kMaxCount);
			}
			return title;
		}

		private static void ExitProcess(Process p)
		{
			try
			{
				p.Kill();
			}
			catch
			{
			}
		}

		private static void ExitAtSame(Process p)
		{
			p.EnableRaisingEvents = true;
			p.Exited += (s, e) =>
			{
				Environment.Exit(0);
			};

			AppDomain.CurrentDomain.ProcessExit += (s, e) =>
			{
				ExitProcess(p);
			};
		}

		[DllImport("user32.dll")]
		private static extern bool EnumThreadWindows(int threadId, EnumThreadDelegate callback, IntPtr lParam);

		private delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

		[DllImport("user32.dll")]
		private static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);

		[DllImport("user32.dll")]
		private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

		private const int SW_HIDE = 0;
		private const int SW_SHOW = 5;

		private static void SwitchWindow(IntPtr hWnd)
		{
			bool success = ShowWindow(hWnd, SW_HIDE);
			if (!success)
			{
				ShowWindow(hWnd, SW_SHOW);
			}
		}

		[DllImport("kernel32.dll")]
		static extern IntPtr GetConsoleWindow();

		[DllImport("kernel32.dll")]
		static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);

		[FlagsAttribute]
		public enum EXECUTION_STATE : uint
		{
			ES_AWAYMODE_REQUIRED = 0x00000040,
			ES_CONTINUOUS = 0x80000000,
			ES_DISPLAY_REQUIRED = 0x00000002,
			ES_SYSTEM_REQUIRED = 0x00000001
		}
	}
}
