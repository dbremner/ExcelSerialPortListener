using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    public class ExcelComms {
        private Excel.Workbook WkBook { get; }
        public string WkSheetName { get; }
        public string RngName { get; }

        [DllImport("Oleacc.dll")]
        static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        [DllImport("User32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildCallback lpEnumFunc, ref IntPtr lParam);

        [DllImport("User32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName( IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public delegate bool EnumChildCallback(IntPtr hwnd, ref IntPtr lParam);

        public ExcelComms(string wkBookName, string wkSheetName, string rngName) {
            if (wkBookName == null) throw new ArgumentNullException(nameof(wkBookName));
            if (wkSheetName == null) throw new ArgumentNullException(nameof(wkSheetName));
            if (rngName == null) throw new ArgumentNullException(nameof(rngName));
            WkBook = WorkbookByName(wkBookName);
            WkSheetName = wkSheetName;
            RngName = rngName;
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
        public Excel.Workbook WorkbookByName(string callingWkbkName) {
            if (callingWkbkName == null) throw new ArgumentNullException(nameof(callingWkbkName));
            foreach (var p in Process.GetProcessesByName("excel")) {
                var winHandle = p.MainWindowHandle;
                //Console.WriteLine($"winHandle = {winHandle}");
                // We need to enumerate the child windows to find one that
                // supports accessibility. To do this, instantiate the
                // delegate and wrap the callback method in it, then call
                // EnumChildWindows, passing the delegate as the 2nd arg.
                if (winHandle != IntPtr.Zero) {
                    var hwndChild = IntPtr.Zero;
                    EnumChildWindows(winHandle, EnumChildProc, ref hwndChild);

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
                        int hr = AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, ref IID_IDispatch, ref ptr);
                        //Console.WriteLine($"hr ptr = {hr}");
                        if (hr >= 0) {
                            // If we successfully got a native OM
                            // IDispatch pointer, we can QI this for
                            // an Excel Application (using the implicit
                            // cast operator supplied in the PIA).
                            var app = ptr.Application;
                            foreach (Excel.Workbook wkbk in app.Workbooks) {
                                if (wkbk.Name == callingWkbkName) {
                                    //Console.WriteLine($"Workbook name = {wkbk.Name}");
                                    return wkbk;
                                }
                            }
                        }
                    }
                }
            }
            //Console.WriteLine($"Failed to find Workbook named '{callingWkbkName}'");
            return null;
        }

        public bool WriteValueToWks(string valueToWrite) {
            if (valueToWrite == null) throw new ArgumentNullException(nameof(valueToWrite));
            try {
                WkBook.Worksheets[WkSheetName].Range[RngName].Value = valueToWrite;
                return true;
            }
            catch (Exception) {
                //Console.WriteLine($"Failed to write value to Excel spreadsheet {WkBook?.Name}.{WkSheetName}.{RngName}, {e.Message}");
                return false;
            }
        }

        public static bool EnumChildProc(IntPtr hwndChild, ref IntPtr lParam) {
            var buf = new StringBuilder(256);
            GetClassName(hwndChild, buf, buf.MaxCapacity);
            if (buf.ToString() == "EXCEL7") {
                lParam = hwndChild;
                return false;
            }
            return true;
        }
    }
}
