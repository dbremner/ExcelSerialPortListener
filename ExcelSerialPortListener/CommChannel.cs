using System;
using System.IO.Ports;
using System.Threading;
using ExcelSerialPortListener.Properties;

namespace ExcelSerialPortListener {
    public class CommChannel {
        private SerialPort CommPort { get; }
        public bool IsOpen => CommPort.IsOpen;
        //public string Response { get; set; } = string.Empty;

    //=== Constructor(s) ===

        public CommChannel() {
            CommPort = new SerialPort() {
                PortName = Settings.Default.PortName,
                BaudRate = Settings.Default.BaudRate,
                DataBits = Settings.Default.DataBits,
                StopBits = Settings.Default.StopBits,        //None, One, OnePointFive, Two
                Parity = Settings.Default.Parity,            //Even, Mark, None, Odd, Space
                ReceivedBytesThreshold = Settings.Default.ReceivedBytesThreshold,
                //Handshake = Handshake.None;
                //RtsEnable = true;
            };
            if(CommPort.IsOpen) CommPort.Close();
            // add listener event handler
            CommPort.DataReceived += SerialDeviceDataReceivedHandler;
        }

        // === Methods ===
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
            Console.WriteLine("got Print command.");
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
