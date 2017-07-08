using System;
using System.Runtime.InteropServices;
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener
{
    public sealed partial class ExcelComms {
        [SuppressUnmanagedCodeSecurity]
        private static class NativeMethods {
            [DllImport("Oleacc.dll", EntryPoint = "AccessibleObjectFromWindow", ExactSpelling = true)]
            internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, [In] ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref Excel.Window ppvObject);

            [DllImport("User32.dll", EntryPoint = "EnumChildWindows", ExactSpelling = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumChildWindows(IntPtr hWndParent, [MarshalAs(UnmanagedType.FunctionPtr)]EnumChildCallback lpEnumFunc, ref IntPtr lParam);

            [return: MarshalAs(UnmanagedType.Bool)]
            internal delegate bool EnumChildCallback(IntPtr hwnd, ref IntPtr lParam);
        }
    }
}
