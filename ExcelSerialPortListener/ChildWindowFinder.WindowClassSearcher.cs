using System;

namespace ExcelSerialPortListener {
    internal sealed partial class ChildWindowFinder {
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
    }
}
