using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    public class Params {
        public string WorkbookName { get; }
        public string WorksheetName { get; }
        public string RangeName { get; }

        public Params(string[] parameters) {
            if (parameters == null) throw new ArgumentNullException(nameof(parameters));
            if (parameters.Length != 3) throw new ArgumentException("Need 3 arguments:", nameof(parameters));
            WorkbookName = parameters[0];
            WorksheetName = parameters[1];
            RangeName = parameters[2];
        }
    }
}
