using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using JetBrains.Annotations;
using Validation;

namespace ExcelSerialPortListener {
    internal static class Utilities {
        private const string ProcessName = "excel";

        [ItemNotNull]
        [NotNull]
        internal static IReadOnlyList<Process> GetExcelInstances() {
            return Process.GetProcessesByName(ProcessName);
        }

        [ContractAnnotation("=> halt")]
        internal static void FatalError([NotNull] [LocalizationRequired] string message) {
            Requires.NotNullOrWhiteSpace(message, nameof(message));

            MessageBox.Show(message, nameof(ExcelSerialPortListener), MessageBoxButtons.OK, MessageBoxIcon.Error);
            Environment.Exit(1);
        }
    }
}
