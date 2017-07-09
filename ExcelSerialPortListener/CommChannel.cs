using System;
using System.IO;
using System.IO.Ports;
using System.Threading;
using ExcelSerialPortListener.Properties;
using JetBrains.Annotations;
using Validation;

namespace ExcelSerialPortListener {
    internal sealed class CommChannel {
        [NotNull] private readonly SerialPort CommPort = new SerialPort {
            PortName = Settings.Default.PortName,
            BaudRate = Settings.Default.BaudRate,
            DataBits = Settings.Default.DataBits,
            StopBits = Settings.Default.StopBits,        //None, One, OnePointFive, Two
            Parity = Settings.Default.Parity,            //Even, Mark, None, Odd, Space
            ReceivedBytesThreshold = Settings.Default.ReceivedBytesThreshold,
            //Handshake = Handshake.None;
            //RtsEnable = true;
        };

        private readonly Action<string> action;

        private bool IsOpen => CommPort.IsOpen;

        public CommChannel() {
            ClosePort();
            CommPort.DataReceived += SerialDeviceDataReceivedHandler;
        }

        public CommChannel(Action<string> action){
            ClosePort();
            CommPort.DataReceived += SerialDeviceDataReceivedHandler;
            this.action = action;
        }

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

        internal void WriteData([NotNull] string dataString) {
            Requires.NotNull(dataString, nameof(dataString));

            Console.WriteLine("got Print command.");
            if (!IsOpen) {
                CommPort.Open();
            }

            CommPort.Write(dataString);
        }

        private void SerialDeviceDataReceivedHandler([NotNull] object sender, [NotNull] SerialDataReceivedEventArgs e) {
            Requires.NotNull(sender, nameof(sender));

            var sp = (SerialPort)sender;
            Program.Response = sp.ReadExisting();
            Console.WriteLine("Received Response: {0}", Program.Response);
        }
    }
}
