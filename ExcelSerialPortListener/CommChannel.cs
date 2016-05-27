using System;
using System.IO.Ports;
using System.Threading;

namespace ExcelSerialPortListener {
    public class CommChannel {
        public string PortName { get; }
        public string BaudRate { get; }
        public string Parity { get; }
        public string DataBits { get; }
        public string StopBits { get; }
        public SerialPort CommPort { get; } = new SerialPort();
        public bool IsOpen => CommPort.IsOpen;
        //public string Response { get; set; } = string.Empty;

    //=== Constructor(s) ===

        public CommChannel(string portName, string baudRate) :
            this(portName, baudRate, "8")
        { }

        private CommChannel(string portName = "COM3", string baudRate = "19200", 
                           string dataBits = "8", string stopBits = "One", string parity = "None") {
            PortName = portName;
            BaudRate = baudRate;
            DataBits = dataBits;
            StopBits = stopBits;        //None, One, OnePointFive, Two
            Parity = parity;            //Even, Mark, None, Odd, Space
            ConfigurePort();
        }

        // === Methods ===
        private void ConfigurePort() {
            if(CommPort.IsOpen) CommPort.Close();
            CommPort.PortName = PortName;
            CommPort.BaudRate = int.Parse(BaudRate);
            CommPort.DataBits = int.Parse(DataBits);
            CommPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), StopBits, ignoreCase: true);
            CommPort.Parity = (Parity)Enum.Parse(typeof(Parity), Parity, ignoreCase: true);
            CommPort.ReceivedBytesThreshold = 11;
            //CommPort.Handshake = Handshake.None;
            //CommPort.RtsEnable = true;
            // add listener event handler
            CommPort.DataReceived += SerialDeviceDataReceivedHandler;
        }

        public void ClosePort() {
            if (IsOpen) CommPort.Close();
        }

        public bool OpenPort() {
            try {
                CommPort.Open();
                return true;
            } catch {
                return false;
            }
        }

        //public string ReadData(double timeOutInSeconds = 30) {
        //    DateTime timeOut = DateTime.Now.AddSeconds(timeOutInSeconds);
        //    bool isTimedOut = false;
        //    do {
        //        if (Response.Length > 0)
        //            break;
        //        Thread.Sleep(200);
        //        isTimedOut = DateTime.Now > timeOut;
        //    } while (!isTimedOut);

        //    if (isTimedOut) {
        //        return "Timed Out";
        //    } else {
        //        return OnlyDigits(Response);
        //    }
        //}

        public void WriteData(string dataString) {
            Console.WriteLine($"got Print command.");
            if (!IsOpen)
                CommPort.Open();
            CommPort.Write(dataString);
        }

        //private string OnlyDigits(string s) {
        //    string onlyDigits = s.Trim();
        //    int indexOfSpaceG = onlyDigits.IndexOf(" g");
        //    if (indexOfSpaceG > 0)
        //        onlyDigits = onlyDigits.Substring(0, indexOfSpaceG);
        //    double tester;
        //    if (double.TryParse(onlyDigits, out tester)) {
        //        return onlyDigits;
        //    } else {
        //        return string.Empty;
        //    }
        //}

        private void SerialDeviceDataReceivedHandler(object sender, SerialDataReceivedEventArgs e) {
            var sp = (SerialPort)sender;
            Program.Response = sp.ReadExisting();
            Console.WriteLine($"Received Response: {Program.Response}");
        }
    }
}
