using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Text;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    public class ExcelComms {
        private readonly Excel.Workbook _workBook;
        public string WorkSheetName { get; }
        public string RangeName { get; }

        public Excel.Workbook WorkBook
        {
            get { return _workBook; }
        }

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

        public ExcelComms(string workBookName, string workSheetName, string rangeName) {
            if (workBookName == null) throw new ArgumentNullException(nameof(workBookName));
            if (workSheetName == null) throw new ArgumentNullException(nameof(workSheetName));
            if (rangeName == null) throw new ArgumentNullException(nameof(rangeName));
            Contract.EndContractBlock();

            bool found = TryFindWorkbookByName(workBookName, out _workBook);
            if (!found) {
                MessageBox.Show("Excel is not running or requested spreadsheet is not open, exiting now",
                    nameof(ExcelSerialPortListener), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
            WorkSheetName = workSheetName;
            RangeName = rangeName;
        }

        [ContractInvariantMethod]
        private void ObjectInvariant() {
            Contract.Invariant(WorkBook != null);
            Contract.Invariant(WorkSheetName != null);
            Contract.Invariant(RangeName != null);
        }

        /// <summary>
        /// A function that returns the Excel.Workbook object that matches the passed Excel workbook file name.
        /// This function is substantially based on open-source code, not authored by me.
        /// However, none of the several sources that had this code clearly claimed original
        /// authorship, though I believe the author is Andrew Whitechapel. 
        /// @https://www.linkedin.com/in/andrew-whitechapel-083b75
        /// </summary>
        /// <param name="callingWkbkName"></param>
        /// <returns>Excel.Workbook</returns>
        public bool TryFindWorkbookByName(string callingWkbkName, out Excel.Workbook target) {
            if (callingWkbkName == null) throw new ArgumentNullException(nameof(callingWkbkName));
            Contract.EndContractBlock();
            var excelInstances = Process.GetProcessesByName("excel");
            if (excelInstances.Length == 0) {
                target = null;
                return false;
            }

            foreach (var p in excelInstances) {
                Contract.Assume(p != null);
                var winHandle = p.MainWindowHandle;
                //Console.WriteLine($"winHandle = {winHandle}");
                // We need to enumerate the child windows to find one that
                // supports accessibility. To do this, instantiate the
                // delegate and wrap the callback method in it, then call
                // EnumChildWindows, passing the delegate as the 2nd arg.
                if (winHandle != IntPtr.Zero) {
                    var hwndChild = IntPtr.Zero;
                    NativeMethods.EnumChildWindows(winHandle, EnumChildProc, ref hwndChild);

                    // If we found an accessible child window, call
                    // AccessibleObjectFromWindow, passing the constant
                    // OBJID_NATIVEOM (defined in winuser.h) and
                    // IID_IDispatch - we want an IDispatch pointer
                    // into the native object model.
                    //Console.WriteLine($"hwndChild = {hwndChild}");
                    if (hwndChild != IntPtr.Zero) {
                        const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        var IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

                        Excel.Window ptr = null;
                        int hr = NativeMethods.AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, ref IID_IDispatch, ref ptr);
                        //Console.WriteLine($"hr ptr = {hr}");
                        if (hr >= 0) {
                            // If we successfully got a native OM
                            // IDispatch pointer, we can QI this for
                            // an Excel Application (using the implicit
                            // cast operator supplied in the PIA).
                            Contract.Assume(ptr != null);
                            var app = ptr.Application;
                            foreach (Excel.Workbook wkbk in app.Workbooks) {
                                if (wkbk.Name == callingWkbkName) {
                                    //Console.WriteLine($"Workbook name = {wkbk.Name}");
                                    target = wkbk;
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            //Console.WriteLine($"Failed to find Workbook named '{callingWkbkName}'");
            target = null;
            return false;
        }

        public bool TryWriteStringToWorksheet(string valueToWrite) {
            if (valueToWrite == null) throw new ArgumentNullException(nameof(valueToWrite));
            Contract.Requires(WorkBook != null);
            Contract.Requires(WorkBook.Worksheets != null);
            Contract.EndContractBlock();
            try {
                WorkBook.Worksheets[WorkSheetName].Range[RangeName].Value = valueToWrite;
                return true;
            }
            catch (Exception) {
                //Console.WriteLine($"Failed to write value to Excel spreadsheet {WorkBook?.Name}.{WorkSheetName}.{RangeName}, {e.Message}");
                return false;
            }
        }

        public static bool EnumChildProc(IntPtr hwndChild, ref IntPtr lParam) {
            Contract.Requires(hwndChild != IntPtr.Zero);
            var className = PInvoke.User32.GetClassName(hwndChild);
            if (className == "EXCEL7") {
                lParam = hwndChild;
                return false;
            }
            return true;
        }
    }
}
