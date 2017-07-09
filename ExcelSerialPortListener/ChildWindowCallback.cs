using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelSerialPortListener {
    [return: MarshalAs(UnmanagedType.Bool)]
    internal delegate bool ChildWindowCallback(IntPtr hwnd, ref IntPtr lParam);
}