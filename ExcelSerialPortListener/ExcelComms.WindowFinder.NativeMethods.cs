using System;
using System.Runtime.InteropServices;
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    internal sealed partial class ExcelComms {
        private sealed partial class WindowFinder {
            [SuppressUnmanagedCodeSecurity]
            private static class NativeMethods {
                [DllImport("Oleacc.dll", EntryPoint = "AccessibleObjectFromWindow", ExactSpelling = true)]
                internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, [In] ref Guid iid, [Out, MarshalAs(UnmanagedType.IUnknown)] out Excel.Window ppvObject);
            }
        }
    }
}