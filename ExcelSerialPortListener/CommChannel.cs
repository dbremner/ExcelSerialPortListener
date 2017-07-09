using System;
using System.Diagnostics.Contracts;
using System.IO;
using System.IO.Ports;
using System.Threading;
using ExcelSerialPortListener.Properties;
using JetBrains.Annotations;

namespace ExcelSerialPortListener {
    public sealed class CommChannel {
        [NotNull] private readonly SerialPort CommPort;

        private bool IsOpen => CommPort.IsOpen;
        //public string Response { get; set; } = string.Empty;

    //=== Constructor(s) ===

        public CommChannel() {
            CommPort = new SerialPort {
                PortName = Settings.Default.PortName,
                BaudRate = Settings.Default.BaudRate,
                DataBits = Settings.Default.DataBits,
                StopBits = Settings.Default.StopBits,        //None, One, OnePointFive, Two
                Parity = Settings.Default.Parity,            //Even, Mark, None, Odd, Space
                ReceivedBytesThreshold = Settings.Default.ReceivedBytesThreshold,
                //Handshake = Handshake.None;
                //RtsEnable = true;
            };
            if(CommPort.IsOpen) {
                CommPort.Close();
            }
            // add listener event handler
            CommPort.DataReceived += SerialDeviceDataReceivedHandler;
        }

        [ContractInvariantMethod]
        private void ObjectInvariant() {
            Contract.Invariant(CommPort != null);
        }

        // === Methods ===
        internal void ClosePort() {
            if (IsOpen) {
                CommPort.Close();
            }
        }

        internal bool OpenPort() {
            bool result = true;
            void HandleException() => result = false;
            try {
                CommPort.Open();
            }
            catch (UnauthorizedAccessException) {
                //There are several possible causes.
                //1. access is denied
                //2. the current process has already opened it
                //3. another process has already opened it
                HandleException();
            }

            catch (ArgumentOutOfRangeException) {
                //One or more of the properties for this instance are invalid
                HandleException();
            }
            catch (IOException) {
                //The port is in an invalid state or setting the port state failed.
                HandleException();
            }
            catch (InvalidOperationException) {
                //The port is already open
                HandleException();
            }
            return result;
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

        internal void WriteData([NotNull] string dataString) {
            if (dataString == null) {
                throw new ArgumentNullException(nameof(dataString));
            }

            Contract.EndContractBlock();
            Console.WriteLine("got Print command.");
            if (!IsOpen) {
                CommPort.Open();
            }

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

        private void SerialDeviceDataReceivedHandler([NotNull] object sender, [NotNull] SerialDataReceivedEventArgs e) {
            if (sender == null) {
                throw new ArgumentNullException(nameof(sender));
            }

            Contract.EndContractBlock();
            var sp = (SerialPort)sender;
            Program.Response = sp.ReadExisting();
            Console.WriteLine("Received Response: {0}", Program.Response);
        }
    }
}
