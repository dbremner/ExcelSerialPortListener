using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    internal static class Utilities {
        private const string ProcessName = "excel";

        internal static IReadOnlyList<Process> GetExcelInstances() {
            return Process.GetProcessesByName(ProcessName);
        }
    }
}
