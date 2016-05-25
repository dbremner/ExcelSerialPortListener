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
        public string CommPort { get; }
        public string Baudrate { get; }

        public Params(string[] parameters) {
            WorkbookName = parameters[0];
            WorksheetName = parameters[1];
            RangeName = parameters[2];
            CommPort = parameters[3];
            Baudrate = parameters[4];
        }
    }
}
