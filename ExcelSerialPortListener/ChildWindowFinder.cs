using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    internal sealed partial class ChildWindowFinder {
        private readonly IntPtr mainWindow;

        private readonly ChildWindowCallback callback;

        public ChildWindowFinder(IntPtr mainWindow, ChildWindowCallback callback) {
            this.mainWindow = mainWindow;
            this.callback = callback;
        }

        public bool TryFindChildWindow(out IntPtr childWindow) {
            //Console.WriteLine($"winHandle = {winHandle}");
            // We need to enumerate the child windows to find one that
            // supports accessibility.
            childWindow = IntPtr.Zero;
            if (mainWindow != IntPtr.Zero) {
                NativeMethods.EnumChildWindows(mainWindow, callback, ref childWindow);
            }
            return childWindow != IntPtr.Zero;
        }
    }
}
