using JetBrains.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    internal interface IExcelComms {
        [ContractAnnotation("=> false, target:null; => true, target:notnull")]
        bool TryFindWorkbookByName([CanBeNull] out Workbook target);

        bool TryWriteStringToWorksheet([NotNull] Workbook workBook, [NotNull] string valueToWrite);
    }
}