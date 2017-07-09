using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JetBrains.Annotations;
using Validation;

namespace ExcelSerialPortListener {
    internal sealed partial class ChildWindowFinder {
        [NotNull] private readonly ChildWindowCallback callback;

        public ChildWindowFinder([NotNull] ChildWindowCallback callback) {
            Requires.NotNull(callback, nameof(callback));

            this.callback = callback;
        }

        public bool TryFindChildWindow(IntPtr mainWindow, out IntPtr childWindow) {
            Requires.NotNull(mainWindow, nameof(mainWindow));

            childWindow = IntPtr.Zero;
            NativeMethods.EnumChildWindows(mainWindow, callback, ref childWindow);
            return childWindow != IntPtr.Zero;
        }

        private sealed class WindowClassSearcher {
            private readonly string targetClassName;

            public WindowClassSearcher(string targetClassName) {
                this.targetClassName = targetClassName;
            }

            public bool EnumChildProc(IntPtr child, ref IntPtr lParam) {
                var className = PInvoke.User32.GetClassName(child);
                if (className == this.targetClassName) {
                    lParam = child;
                    return false;
                }
                return true;
            }
        }

        public static ChildWindowFinder FindWindowClass(string className) {
            var searcher = new WindowClassSearcher(className);
            return new ChildWindowFinder(searcher.EnumChildProc);
        }
    }
}
