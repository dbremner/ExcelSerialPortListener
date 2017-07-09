using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using JetBrains.Annotations;
using PInvoke;
using Validation;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    public sealed partial class ExcelComms {
        private readonly Excel.Workbook _workBook;
        private const string iidDispatchGuid = "{00020400-0000-0000-C000-000000000046}";
        private Guid IID_IDispatch = new Guid(iidDispatchGuid);

        [NotNull]
        private string WorkSheetName { get; }

        [NotNull]
        private string RangeName { get; }

        private Excel.Workbook WorkBook
        {
            get { return _workBook; }
        }

        public ExcelComms([NotNull] string workBookName, [NotNull] string workSheetName, [NotNull] string rangeName) {
            Requires.NotNullOrWhiteSpace(workBookName, nameof(workBookName));
            Requires.NotNullOrWhiteSpace(workSheetName, nameof(workSheetName));
            Requires.NotNullOrWhiteSpace(rangeName, nameof(rangeName));

            if (!TryFindWorkbookByName(workBookName, out _workBook)) {
                MessageBox.Show("Excel is not running or requested spreadsheet is not open, exiting now",
                    nameof(ExcelSerialPortListener), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
            (WorkSheetName, RangeName) = (workSheetName, rangeName);
        }

        /// <summary>
        /// A function that returns the Excel.Workbook object that matches the passed Excel workbook file name.
        /// This function is substantially based on open-source code, not authored by me.
        /// However, none of the several sources that had this code clearly claimed original
        /// authorship, though I believe the author is Andrew Whitechapel. 
        /// @https://www.linkedin.com/in/andrew-whitechapel-083b75
        /// </summary>
        /// <param name="callingWkbkName"></param>
        /// <param name="target"></param>
        /// <returns>Excel.Workbook</returns>
        private bool TryFindWorkbookByName([NotNull] string callingWkbkName, out Excel.Workbook target) {
            Requires.NotNullOrWhiteSpace(callingWkbkName, nameof(callingWkbkName));

            IReadOnlyList<Process> excelInstances = Process.GetProcessesByName("excel");
            if (excelInstances.Count == 0) {
                target = null;
                return false;
            }

            foreach (var p in excelInstances) {
                var winHandle = p.MainWindowHandle;
                if (winHandle == IntPtr.Zero) {
                    continue;
                }
                //Console.WriteLine($"hwndChild = {hwndChild}");
                if (!TryFindAccessibleChildWindow(winHandle, out var hwndChild)) {
                    continue;
                }
                //Console.WriteLine($"hr ptr = {hr}");
                if (!TryGetExcelWindow(hwndChild, out Excel.Window ptr)) {
                    continue;
                }
                // If we successfully got a native OM
                // IDispatch pointer, we can QI this for
                // an Excel Application (using the implicit
                // cast operator supplied in the PIA).
                var app = ptr.Application;
                foreach (Excel.Workbook wkbk in app.Workbooks) {
                    if (wkbk.Name == callingWkbkName) {
                        //Console.WriteLine($"Workbook name = {wkbk.Name}");
                        target = wkbk;
                        return true;
                    }
                }
            }
            //Console.WriteLine($"Failed to find Workbook named '{callingWkbkName}'");
            target = null;
            return false;
        }

        private static bool TryFindAccessibleChildWindow(IntPtr mainWindow, out IntPtr childWindow) {
            //Console.WriteLine($"winHandle = {winHandle}");
            // We need to enumerate the child windows to find one that
            // supports accessibility. To do this, instantiate the
            // delegate and wrap the callback method in it, then call
            // EnumChildWindows, passing the delegate as the 2nd arg.
            if (mainWindow != IntPtr.Zero) {
                var hwndChild = IntPtr.Zero;

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

                NativeMethods.EnumChildWindows(mainWindow, EnumChildProc, ref hwndChild);
                childWindow = hwndChild;
            }
            else {
                childWindow = IntPtr.Zero;
            }
            return childWindow != IntPtr.Zero;
        }

        private bool TryGetExcelWindow(IntPtr hwndChild, out Excel.Window ptr) {
            // If we found an accessible child window, call
            // AccessibleObjectFromWindow, passing the constant
            // OBJID_NATIVEOM (defined in winuser.h) and
            // IID_IDispatch - we want an IDispatch pointer
            // into the native object model.
            const uint OBJID_NATIVEOM = 0xFFFFFFF0;

            HResult hr = NativeMethods.AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, ref IID_IDispatch, out ptr);
            return hr.Succeeded;
        }

        internal bool TryWriteStringToWorksheet([NotNull] string valueToWrite) {
            Requires.NotNullOrWhiteSpace(valueToWrite, nameof(valueToWrite));
            Requires.NotNull(WorkBook, nameof(WorkBook));
            Requires.NotNull(WorkBook.Worksheets, nameof(WorkBook.Worksheets));

            try {
                WorkBook.Worksheets[WorkSheetName].Range[RangeName].Value = valueToWrite;
                return true;
            }
            catch (Exception) {
                //Console.WriteLine($"Failed to write value to Excel spreadsheet {WorkBook?.Name}.{WorkSheetName}.{RangeName}, {e.Message}");
                return false;
            }
        }
    }
}
