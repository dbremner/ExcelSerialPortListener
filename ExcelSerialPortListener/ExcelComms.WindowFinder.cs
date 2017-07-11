using System;
using JetBrains.Annotations;
using PInvoke;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    internal sealed partial class ExcelComms {
        private sealed partial class WindowFinder {
            private const string IidDispatchGuid = "{00020400-0000-0000-C000-000000000046}";
            private Guid IID_IDispatch = new Guid(IidDispatchGuid);

            [ContractAnnotation("=> false, ptr:null; => true, ptr:notnull")]
            public bool TryFindExcelWindow(IntPtr hwndChild, [CanBeNull] out Excel.Window ptr) {
                // If we found an accessible child window, call
                // AccessibleObjectFromWindow, passing the constant
                // OBJID_NATIVEOM (defined in winuser.h) and
                // IID_IDispatch - we want an IDispatch pointer
                // into the native object model.
                const uint OBJID_NATIVEOM = 0xFFFFFFF0;

                HResult hr = NativeMethods.AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, ref IID_IDispatch, out ptr);
                return hr.Succeeded;
            }
        }
    }
}
