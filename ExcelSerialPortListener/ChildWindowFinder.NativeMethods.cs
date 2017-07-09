using System;
using System.Runtime.InteropServices;
using System.Security;

namespace ExcelSerialPortListener
{
    internal sealed partial class ChildWindowFinder
    {
        [SuppressUnmanagedCodeSecurity]
        private static class NativeMethods
        {
            [DllImport("User32.dll", EntryPoint = "EnumChildWindows", ExactSpelling = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumChildWindows(IntPtr hWndParent, [MarshalAs(UnmanagedType.FunctionPtr)]ChildWindowCallback lpEnumFunc, ref IntPtr lParam);
        }
    }
}
