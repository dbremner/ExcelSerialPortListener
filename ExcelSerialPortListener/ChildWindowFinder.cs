using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener
{
    internal sealed class ChildWindowFinder
    {
        [SuppressUnmanagedCodeSecurity]
        private static class NativeMethods
        {
            [DllImport("User32.dll", EntryPoint = "EnumChildWindows", ExactSpelling = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumChildWindows(IntPtr hWndParent, [MarshalAs(UnmanagedType.FunctionPtr)]ChildWindowCallback lpEnumFunc, ref IntPtr lParam);
        }

        private readonly IntPtr mainWindow;

        private readonly ChildWindowCallback callback;

        public ChildWindowFinder(IntPtr mainWindow, ChildWindowCallback callback)
        {
            this.mainWindow = mainWindow;
            this.callback = callback;
        }

        public static bool TryFindAccessibleChildWindow(IntPtr mainWindow, out IntPtr childWindow) {
            //Console.WriteLine($"winHandle = {winHandle}");
            // We need to enumerate the child windows to find one that
            // supports accessibility.
            childWindow = IntPtr.Zero;
            if (mainWindow != IntPtr.Zero) {

                bool EnumChildProc(IntPtr child, ref IntPtr lParam)
                {
                    var className = PInvoke.User32.GetClassName(child);
                    if (className == "EXCEL7")
                    {
                        lParam = child;
                        return false;
                    }
                    return true;
                }

                NativeMethods.EnumChildWindows(mainWindow, EnumChildProc, ref childWindow);
            }
            return childWindow != IntPtr.Zero;
        }
    }
}
