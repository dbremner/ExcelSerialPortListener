﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using JetBrains.Annotations;
using Validation;
using static ExcelSerialPortListener.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    internal sealed partial class ExcelComms {
        private readonly Excel.Workbook _workBook;
        [NotNull]
        private readonly ChildWindowFinder childWindowFinder = ChildWindowFinder.FindWindowClass("EXCEL7");

        [NotNull]
        private readonly WindowFinder windowFinder = new WindowFinder();

        [NotNull]
        private string WorkSheetName { get; }

        [NotNull]
        private string RangeName { get; }

        private Excel.Workbook WorkBook => _workBook;

        public ExcelComms([NotNull] string workBookName, [NotNull] string workSheetName, [NotNull] string rangeName) {
            Requires.NotNullOrWhiteSpace(workBookName, nameof(workBookName));
            Requires.NotNullOrWhiteSpace(workSheetName, nameof(workSheetName));
            Requires.NotNullOrWhiteSpace(rangeName, nameof(rangeName));

            if (!TryFindWorkbookByName(workBookName, out _workBook)) {
                ErrorMessage("Excel is not running or requested spreadsheet is not open, exiting now");
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

            var excelInstances = GetExcelInstances();
            if (excelInstances.Count == 0) {
                target = null;
                return false;
            }

            foreach (var p in excelInstances) {
                var winHandle = p.MainWindowHandle;
                if (winHandle == IntPtr.Zero) {
                    continue;
                }
                if (!childWindowFinder.TryFindChildWindow(winHandle, out var hwndChild)) {
                    continue;
                }
                if (!windowFinder.TryGetExcelWindow(hwndChild, out Excel.Window ptr)) {
                    continue;
                }
                // If we successfully got a native OM
                // IDispatch pointer, we can QI this for
                // an Excel Application (using the implicit
                // cast operator supplied in the PIA).
                var workbooks = ptr.Application.Workbooks;
                foreach (Excel.Workbook wkbk in workbooks) {
                    if (wkbk.Name == callingWkbkName) {
                        target = wkbk;
                        return true;
                    }
                }
            }
            target = null;
            return false;
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
                return false;
            }
        }
    }
}
