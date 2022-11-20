using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WXLessonRecorder.CLI
{
	public class ScreenShots : IDisposable
	{
		private Bitmap lastScreenShot = null;
		public Bitmap LastScreenShot { get { return lastScreenShot; } }

		//http://www.voidcn.com/article/p-ycvuxmrq-byd.html -> https://stackoverflow.com/questions/10233055/how-to-get-screenshot-to-include-the-invoking-window-on-xp/10234693
		public Bitmap ScreenShotDefault(Screen screen)
		{
			Size sz = screen.Bounds.Size;
			IntPtr hDesk = GetDesktopWindow();
			IntPtr hSrce = GetWindowDC(hDesk);
			IntPtr hDest = CreateCompatibleDC(hSrce);
			IntPtr hBmp = CreateCompatibleBitmap(hSrce, sz.Width, sz.Height);
			IntPtr hOldBmp = SelectObject(hDest, hBmp);
			bool b = BitBlt(hDest, 0, 0, sz.Width, sz.Height, hSrce, screen.Bounds.X, screen.Bounds.Y, CopyPixelOperation.SourceCopy | CopyPixelOperation.CaptureBlt);
			Bitmap bmp = Image.FromHbitmap(hBmp);
			SelectObject(hDest, hOldBmp);
			DeleteObject(hBmp);
			DeleteDC(hDest);
			ReleaseDC(hDesk, hSrce);

			lastScreenShot?.Dispose();
			lastScreenShot = bmp;
			return bmp;
		}
		[DllImport("gdi32.dll")]
		static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int
		wDest, int hDest, IntPtr hdcSource, int xSrc, int ySrc, CopyPixelOperation rop);
		[DllImport("user32.dll")]
		static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDc);
		[DllImport("gdi32.dll")]
		static extern IntPtr DeleteDC(IntPtr hDc);
		[DllImport("gdi32.dll")]
		static extern IntPtr DeleteObject(IntPtr hDc);
		[DllImport("gdi32.dll")]
		static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
		[DllImport("gdi32.dll")]
		static extern IntPtr CreateCompatibleDC(IntPtr hdc);
		[DllImport("gdi32.dll")]
		static extern IntPtr SelectObject(IntPtr hdc, IntPtr bmp);
		[DllImport("user32.dll")]
		static extern IntPtr GetDesktopWindow();
		[DllImport("user32.dll")]
		static extern IntPtr GetWindowDC(IntPtr ptr);

		public void Clear()
		{
			lastScreenShot?.Dispose();
			lastScreenShot = null;
		}

		public static Screen[] GetScreens()
		{
			return Screen.AllScreens;
		}

		#region ===IDisposable===
		private bool disposed = false;

		public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources.
                }

                // unmanaged resources here.
                Clear();

                disposed = true;
            }
        }

        ~ScreenShots()
        {
            Dispose(disposing: false);
        }
        #endregion
    }
}
