using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JetBrains.Annotations;

namespace ExcelSerialPortListener {
    internal static class Utilities {
        private const string ProcessName = "excel";

        [NotNull]
        internal static IReadOnlyList<Process> GetExcelInstances() {
            return Process.GetProcessesByName(ProcessName);
        }
    }
}
