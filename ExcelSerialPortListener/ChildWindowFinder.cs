using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JetBrains.Annotations;
using Validation;

namespace ExcelSerialPortListener {
    internal sealed partial class ChildWindowFinder {
        private readonly IntPtr mainWindow;

        [NotNull] private readonly ChildWindowCallback callback;

        public ChildWindowFinder(IntPtr mainWindow, [NotNull] ChildWindowCallback callback) {
            Requires.NotNull(mainWindow, nameof(mainWindow));
            Requires.NotNull(callback, nameof(callback));
            this.mainWindow = mainWindow;
            this.callback = callback;
        }

        public bool TryFindChildWindow(out IntPtr childWindow) {
            childWindow = IntPtr.Zero;
            NativeMethods.EnumChildWindows(mainWindow, callback, ref childWindow);
            return childWindow != IntPtr.Zero;
        }
    }
}
