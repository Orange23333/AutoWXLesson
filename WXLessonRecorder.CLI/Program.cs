using OpenCvSharp;
using Serilog;
using Log = Serilog.Log;

namespace WXLessonRecorder.CLI
{
	internal class Program
	{
		#region ===可修改变量===
		public static string WindowTitle_WXGroup = "高复15班·班级群";
		public static string WindowTitle_WXLive = "企业微信直播";

		public static double ScanThreshold = 0.9;

		public static string Path_OBSExe = "C:/Program Files/obs-studio/bin/64bit/obs64.exe";

		public static string KeyRec_StartRecord = "^+B";
		public static string KeyRec_StopRecord = "^+E";
		public static int MouseWheelDelta_Normal = -300;
		public static int MouseWheelDelta_AfterClass = -2100;

		//public static int Timeout_OpenClass_ms = 60 * 1000; //1min
		//public static int Retry_OpenClass_time = 3;
		//public static int Timeout_OpenPlayback_ms = 60 * 1000; //1min
		//public static int Retry_OpenPlayback_time = 3;
		//public static int Timeout_WaitDownClass_ms = 10 * 1000; //10s
		//public static int Timeout_SearchNewClass_ms = 10 * 60 * 1000; //10min
		//public static int Delay_WaitRecordToSave_ms = 3 * 1000; //3s

		public static int Timeout_OpenClass_ms = 60 * 1000; //1min
		public static int Retry_OpenClass_time = 3;
		public static int Timeout_OpenPlayback_ms = 60 * 1000; //1min
		public static int Retry_OpenPlayback_time = 3;
		public static int Timeout_WaitDownClass_ms = 60 * 60 * 1000; //1h
		public static int Timeout_SearchNewClass_ms = 10 * 60 * 1000; //10min
		public static int Delay_WaitRecordToSave_ms = 60 * 1000; //1min
		#endregion

		#region ===一般不修改的常量===
		public static readonly string ImagePath_ClassBeginNotification = "images/ClassBeginNotification.png";
		public static readonly string ImagePath_WatchPlaybackButton = "images/WatchPlaybackButton.png";
		public static readonly string ImagePath_WXMicroProgram_CloseButton = "images/WXMicroProgram.CloseButton.png";

		public static readonly string ImagePath_ScreenShotCache = "cache/screenshot.png";
		#endregion

		private static bool afterClassRolling = false;

		static void Main(string[] args)
		{
			Log.Logger = new LoggerConfiguration()
				.MinimumLevel.Verbose()
				.WriteTo.Console()
				.WriteTo.File(
					"log.txt",
					rollingInterval: RollingInterval.Infinite,
					outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}"
				)
				.CreateLogger();

			#region ===初始化===
			Mat mat_ClassBeginNotification = new Mat(ImagePath_ClassBeginNotification);
			Mat mat_WatchPlaybackButton = new Mat(ImagePath_WatchPlaybackButton);
			Mat mat_WXMicroProgram_CloseButton = new Mat(ImagePath_WXMicroProgram_CloseButton);
			#endregion

			Log.Information("请打开微信群窗口，并将其独立出来，并保证一定的大小（至少能显示一个上课通知）。接着滑动到最开始需要录制的消息位置。记得打开OBS。按下任意键开始……");
			Console.ReadKey();

			Log.Information("在命令窗口中按下Ctrl+C结束（直接Alt+F4，用任务管理器也行）。");
			Console.CancelKeyPress += (object? sender, ConsoleCancelEventArgs e) =>
			{
				Exit(-1);
			};

			DateTime timer=DateTime.Now;
			while (true)
			{
				var r = FindClassBeginNotification();
				if (r == AutoEventReturn.Pass)
				{
					timer = DateTime.Now;
					afterClassRolling = true;
				}
				else if(r == AutoEventReturn.Exit)
				{
					break;
				}
				if (IsTimePass(timer, Timeout_SearchNewClass_ms))
				{
					Log.Error("！寻找新课超时。（也许已经没有新课了）");
					break;
				}
			}

			Log.Information("按下任意键退出……");
			Console.ReadKey();
			Exit(0);
		}
		public static void Exit(int exitCode)
		{
			Log.CloseAndFlush();
			Environment.Exit(exitCode);
		}

		public static bool IsTimePass(DateTime begin, int ms)
		{
			return DateTime.Now.Subtract(begin) >= new TimeSpan(0, 0, 0, 0, ms);
		}

		public enum AutoEventReturn
		{
			Fail=-1,
			Exit=0,
			Pass=1
		}

		private static AutoEventReturn FindClassBeginNotification()
		{
			Log.Information("第一步：寻找上课通知。");

			IntPtr wxGroupWindowHandle = User32.FindWindow(null, WindowTitle_WXGroup);
			if (wxGroupWindowHandle == IntPtr.Zero)
			{
				Log.Error($"未检测到微信群窗口“{WindowTitle_WXGroup}”。");
				Thread.Sleep(1000);
				return AutoEventReturn.Fail;
			}
			Log.Information(String.Format("获取到微信群窗口“{0}”，窗口句柄为{1:X}。", WindowTitle_WXGroup, (int)wxGroupWindowHandle));

			User32.SetWindowPos(wxGroupWindowHandle, User32.HWND_TOPMOST, 0, 0, 0, 0, User32.SWP_NOSIZE | User32.SWP_NOMOVE | User32.SWP_SHOWWINDOW | User32.SWP_ASYNCWINDOWPOS);
			Log.Verbose("置顶微信群窗口。");

			User32.GetWindowRect(wxGroupWindowHandle, out Rectangle clientRect);
			//Cursor.Position = new System.Drawing.Point(clientRect.Left + clientRect.Width / 2, clientRect.Top + clientRect.Height / 2);
			Cursor.Position = new System.Drawing.Point(clientRect.Left + 8, clientRect.Top + 80);
			Log.Verbose($"移动鼠标至矩形“{clientRect}”内。");

			SimClick(Cursor.Position.X, Cursor.Position.Y);

			if (afterClassRolling)
			{
				User32.mouse_event(User32.MOUSEEVENTF_WHEEL, 0, 0, MouseWheelDelta_AfterClass, 0);
				Log.Verbose($"滚动鼠标（d={MouseWheelDelta_AfterClass}）。");
				afterClassRolling = false;
			}

			User32.mouse_event(User32.MOUSEEVENTF_WHEEL, 0, 0, MouseWheelDelta_Normal, 0);
			Log.Verbose($"滚动鼠标（d={MouseWheelDelta_Normal}）。");
			Thread.Sleep(1500);

			return OpenClass();
		}
		private static AutoEventReturn OpenClass()
		{
			Log.Information("第一步：打开课堂信息页面。");
			if (SearchPosition(ImagePath_ClassBeginNotification, out System.Drawing.Point pos))
			{
				Log.Information("扫描到课堂信息。");
				for (int i = 1; i <= Retry_OpenClass_time; i++)
				{
					Log.Debug($"[第{i}次尝试点击]打开课堂界面。");
					SimClick(pos.X, pos.Y);

					var r = OpenPlayback();
					if (r != AutoEventReturn.Fail)
					{
						return r;
					}
				}
			}
			else
			{
				Log.Debug("未扫描到课堂信息。");
			}
			return AutoEventReturn.Fail;
		}
		private static AutoEventReturn OpenPlayback()
		{
			Log.Information("第三步：寻找回放按钮。");

			DateTime timer = DateTime.Now;
			IntPtr wxLiveWindowHandle;
			while (true)
			{
				wxLiveWindowHandle = User32.FindWindow(null, WindowTitle_WXLive);
				if (wxLiveWindowHandle == IntPtr.Zero)
				{
					Log.Debug($"未检测到企业微信直播“{WindowTitle_WXLive}”。");
				}
				else
				{
					Log.Information(String.Format("获取到企业微信直播窗口“{0}”，窗口句柄为{1:X}。", WindowTitle_WXLive, (int)wxLiveWindowHandle));
					while (true)
					{
						User32.SetWindowPos(wxLiveWindowHandle, User32.HWND_TOPMOST, 0, 0, 0, 0, User32.SWP_NOSIZE | User32.SWP_NOMOVE | User32.SWP_SHOWWINDOW | User32.SWP_ASYNCWINDOWPOS);
						Log.Verbose("置顶微信群窗口。");

						Log.Debug("尝试扫描回放按钮。");
						if (SearchPosition(ImagePath_WatchPlaybackButton, out System.Drawing.Point pos))
						{
							Log.Information("扫描到回放按钮。");
							Log.Information("第四步：打开回放。");

							for (int i = 0; i < Retry_OpenPlayback_time; i++)
							{
								SimClick(pos.X, pos.Y);

								if (!SearchPosition(ImagePath_WatchPlaybackButton, out System.Drawing.Point pos2))
								{
									return BeginRecord();
								}
							}
							Log.Error("！打开回放失败。");
							return AutoEventReturn.Fail;
						}
						if (IsTimePass(timer, Timeout_OpenPlayback_ms))
						{
							Log.Error("！寻找回放超时。（也许已经没有回放了）");
							return AutoEventReturn.Fail;
						}
						Thread.Sleep(1000);
					}
				}
				if (IsTimePass(timer, Timeout_OpenPlayback_ms))
				{
					Log.Error("！打开直播窗口超时。");
					return AutoEventReturn.Fail;
				}
				Thread.Sleep(1000);
			}
		}
		private static AutoEventReturn BeginRecord()
		{
			//System.Windows.Forms.SendKeys.SendWait(KeyRec_StartRecord);
			User32.keybd_event(User32.vbKeyControl, 0, 0, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyShift, 0, 0, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyB, 0, 0, 0);
			Thread.Sleep(100);
			User32.keybd_event(User32.vbKeyControl, 0, 2, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyShift, 0, 2, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyB, 0, 2, 0);
			Log.Information("开始录制。");
			
			Thread.Sleep(Timeout_WaitDownClass_ms);

			//System.Windows.Forms.SendKeys.SendWait(KeyRec_StopRecord);
			User32.keybd_event(User32.vbKeyControl, 0, 0, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyShift, 0, 0, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyE, 0, 0, 0);
			Thread.Sleep(100);
			User32.keybd_event(User32.vbKeyControl, 0, 2, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyShift, 0, 2, 0);
			Thread.Sleep(50);
			User32.keybd_event(User32.vbKeyE, 0, 2, 0);
			Log.Information("结束录制（开始等待录制输出结束）。");

			Thread.Sleep(Delay_WaitRecordToSave_ms);
			Log.Information("录制完成。");

			return AutoEventReturn.Pass;
		}

		public static bool SearchPosition(string templatePath, out System.Drawing.Point pos)
		{
			ScreenShots screenShots = new ScreenShots();
			Mat screenShot = null;
			Mat template = null;
			Mat result = null;
			Screen screen = Screen.AllScreens[0];

			//获取并保存截屏
			screenShots.ScreenShotDefault(screen);
			screenShots.LastScreenShot.Save(ImagePath_ScreenShotCache);
			//初始化图片
			screenShot = new Mat(ImagePath_ScreenShotCache);
			template = new Mat(templatePath);
			result = new Mat();//Mat(new OpenCvSharp.Size(screenShots.LastScreenShot.Width, screenShots.LastScreenShot.Height), MatType.CV_32SC1)

			//比对
			Cv2.MatchTemplate(InputArray.Create(screenShot), InputArray.Create(template), OutputArray.Create(result), TemplateMatchModes.CCoeffNormed, null);
			//查找最佳匹配
			OpenCvSharp.Point maxLoc = new OpenCvSharp.Point(0, 0);
			double maxVal = 0;
			Cv2.MinMaxLoc(InputArray.Create(result), out _, out maxVal, out _, out maxLoc);

			if (maxVal >= ScanThreshold)
			{
				//计算实际中心位置
				int realX = screen.WorkingArea.X + maxLoc.X + template.Width / 2;
				int realY = screen.WorkingArea.Y + maxLoc.Y + template.Height / 2;

				pos = new System.Drawing.Point(realX, realY);
				Log.Verbose($"找到目标中心点({realX},{realY})。");
				return true;
			}
			else
			{
				pos = System.Drawing.Point.Empty;
				Log.Verbose($"未找到目标。");
				return false;
			}
		}

		public static void SimClick(int x, int y)
		{
			Cursor.Position = new System.Drawing.Point(x, y);
			Thread.Sleep(10);
			User32.mouse_event(User32.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
			Thread.Sleep(10);
			User32.mouse_event(User32.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
			Thread.Sleep(10);
			Log.Verbose($"SimClick({x},{y})");
		}
	}
}