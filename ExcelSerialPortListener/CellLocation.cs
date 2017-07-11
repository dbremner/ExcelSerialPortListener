using JetBrains.Annotations;
using Validation;

namespace ExcelSerialPortListener {
    internal sealed class CellLocation {
        [NotNull]
        public string WorkBookName { get; }

        [NotNull]
        public string WorkSheetName { get; }

        [NotNull]
        public string RangeName { get; }

        public CellLocation([NotNull] string workBookName, [NotNull] string workSheetName, [NotNull] string rangeName) {
            Requires.NotNullOrWhiteSpace(workBookName, nameof(workBookName));
            Requires.NotNullOrWhiteSpace(workSheetName, nameof(workSheetName));
            Requires.NotNullOrWhiteSpace(rangeName, nameof(rangeName));
            this.WorkBookName = workBookName;
            this.WorkSheetName = workSheetName;
            this.RangeName = rangeName;
        }

        public void Deconstruct(
            [NotNull] out string workBookName,
            [NotNull] out string workSheetName,
            [NotNull] out string rangeName) {
            (workBookName, workSheetName, rangeName) = (WorkBookName, WorkSheetName, RangeName);
        }
    }
}
