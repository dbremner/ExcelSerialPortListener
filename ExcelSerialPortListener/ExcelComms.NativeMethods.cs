using System;
using System.Runtime.InteropServices;
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    public sealed partial class ExcelComms {
        [SuppressUnmanagedCodeSecurity]
        internal static class NativeMethods {
            [DllImport("Oleacc.dll", EntryPoint = "AccessibleObjectFromWindow", ExactSpelling = true)]
            internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, [In] ref Guid iid, [Out, MarshalAs(UnmanagedType.IUnknown)] out Excel.Window ppvObject);
        }
    }
}
