using System;
using System.Windows.Forms;

namespace EODAddIn.Utils
{
    public class WinHwnd : IWin32Window
    {
        public IntPtr Handle => (IntPtr)Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
    }
}
