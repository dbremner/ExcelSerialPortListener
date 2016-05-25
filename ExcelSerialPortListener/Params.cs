using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    public struct Params {
        public string WorkbookName;
        public string WorksheetName;
        public string RangeName;
        public string CommPort;
        public string Baudrate;

        public Params(string[] parameters) {
            WorkbookName = parameters[0];
            WorksheetName = parameters[1];
            RangeName = parameters[2];
            CommPort = parameters[3];
            Baudrate = parameters[4];
        }
    }
}
