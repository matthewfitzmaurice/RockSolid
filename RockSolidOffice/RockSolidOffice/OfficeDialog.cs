using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using log4net;

namespace RockSolidOffice
{
    public class OfficeDialog : Window
    {
        static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType); //See http://logging.apache.org/log4net/index.html

        public OfficeDialog()
        {
            try
            {
                if (log.IsInfoEnabled) log.Info(System.Reflection.MethodBase.GetCurrentMethod().Name);
                using (Process currentProcess = Process.GetCurrentProcess())
                    SetCentering(this, currentProcess.MainWindowHandle);
            }
            catch (Exception ex)
            {
                log.Error(System.Reflection.MethodBase.GetCurrentMethod().Name, ex);
            }
        }

        static void SetCentering(Window win, IntPtr ownerHandle)
        {
            if (log.IsInfoEnabled) log.InfoFormat("{0} {1}", System.Reflection.MethodBase.GetCurrentMethod().Name, ownerHandle);
            bool isWindow = IsWindow(ownerHandle);
            if (!isWindow) //Don't try and centre the window if the ownerHandle is invalid.  To resolve Poyner issue with invalid window handle error
            {
                log.InfoFormat("ownerHandle IsWindow: {0}", isWindow);
                return;
            }
            //Show in center of owner if win form.
            if (ownerHandle.ToInt32() != 0)
            {
                var helper = new WindowInteropHelper(win);
                helper.Owner = ownerHandle;
                win.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
                win.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        }

        //protected override void OnSourceInitialized(EventArgs e)
        //{
        //    base.OnSourceInitialized(e);
        //    WindowHelper.RemoveIcon(this);
        //}

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindow(IntPtr hWnd);
    }
}
