using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WXLessonRecorder.CLI
{
	public class WeiXinLessonWindow
	{
		public static Rectangle[] Locate()
		{
			;
		}
	
		public static WeiXinLessonInfo GetInfo()
		{
			;
		}
	
		public static Bitmap GetWeiXinLessonWindowScreenshot()
		{
			ScreenShots screenShots = new ScreenShots();
			Rectangle location = Locate()[0];
			Screen screen = ScreenShots.GetScreens()[0];
	
			screenShots.ScreenShotDefault(screen);
			Graphics g = Graphics.FromImage(screenShots.LastScreenShot);
			g;
	
			Image.FromStream(g.)
	
			return g.To;
		}
	}

	public class WeiXinLessonInfo
	{
		public string Title;
		public bool IsLiving;
		public DateTime? Begin;
		public DateTime? End;
		public bool? HasPlayback;
	}
}
