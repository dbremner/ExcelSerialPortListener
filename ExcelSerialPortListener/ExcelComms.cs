using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using JetBrains.Annotations;
using Validation;
using static ExcelSerialPortListener.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

// ReSharper disable HeapView.ObjectAllocation.Possible
namespace ExcelSerialPortListener {
    internal sealed partial class ExcelComms : IExcelComms {
        [NotNull]
        private readonly ChildWindowFinder childWindowFinder = ChildWindowFinder.FindWindowClass("EXCEL7");

        [NotNull]
        private readonly WindowFinder windowFinder = new WindowFinder();

        [NotNull]
        private readonly CellLocation cellLocation;

        public ExcelComms([NotNull] CellLocation cellLocation) {
            Requires.NotNull(cellLocation, nameof(cellLocation));

            this.cellLocation = cellLocation;
        }

        /// <summary>
        /// A function that returns the Excel.Workbook object that matches the passed Excel workbook file name.
        /// This function is substantially based on open-source code, not authored by me.
        /// However, none of the several sources that had this code clearly claimed original
        /// authorship, though I believe the author is Andrew Whitechapel.
        /// @https://www.linkedin.com/in/andrew-whitechapel-083b75
        /// </summary>
        /// <param name="target">matching workbook or null</param>
        /// <returns>Excel.Workbook</returns>
        [ContractAnnotation("=> false, target:null; => true, target:notnull")]
        public bool TryFindWorkbookByName(out Excel.Workbook target) {
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

                if (!windowFinder.TryFindExcelWindow(hwndChild, out Excel.Window ptr)) {
                    continue;
                }

                // If we successfully got a native OM
                // IDispatch pointer, we can QI this for
                // an Excel Application (using the implicit
                // cast operator supplied in the PIA).
                var workbooks = ptr.Application.Workbooks;
                if (TryFindWorkbook(workbooks, out var victim)) {
                    target = victim;
                    return true;
                }
            }

            target = null;
            return false;
        }

        public bool TryWriteStringToWorksheet(Excel.Workbook workBook, string valueToWrite) {
            Requires.NotNullOrWhiteSpace(valueToWrite, nameof(valueToWrite));
            Requires.NotNull(workBook, nameof(workBook));
            Requires.NotNull(workBook.Worksheets, nameof(workBook.Worksheets));

            try {
                var(_, workSheetName, rangeName) = cellLocation;
                workBook.Worksheets[workSheetName].Range[rangeName].Value = valueToWrite;
                return true;
            }
            catch (Exception) {
                return false;
            }
        }

        [ContractAnnotation("=> false, target:null; => true, target:notnull")]
        private bool TryFindWorkbook([NotNull] Excel.Workbooks workbooks, [CanBeNull] out Excel.Workbook target) {
            Requires.NotNull(workbooks, nameof(workbooks));

            foreach (Excel.Workbook workbook in workbooks) {
                if (workbook.Name == cellLocation.WorkBookName) {
                    target = workbook;
                    return true;
                }
            }

            target = null;
            return false;
        }
    }
}