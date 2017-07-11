using System;
using System.IO;
using System.IO.Ports;
using System.Threading;
using ExcelSerialPortListener.Properties;
using JetBrains.Annotations;
using Validation;

// ReSharper disable HeapView.ClosureAllocation
namespace ExcelSerialPortListener {
    internal sealed class CommChannel : IDisposable {
        [NotNull]
        private readonly SerialPort commPort = new SerialPort {
            PortName = Settings.Default.PortName,
            BaudRate = Settings.Default.BaudRate,
            DataBits = Settings.Default.DataBits,
            StopBits = Settings.Default.StopBits,        // None, One, OnePointFive, Two
            Parity = Settings.Default.Parity,            // Even, Mark, None, Odd, Space
            ReceivedBytesThreshold = Settings.Default.ReceivedBytesThreshold,

            // Handshake = Handshake.None;
            // RtsEnable = true;
        };

        [NotNull]
        private readonly Action<string> action;

        public CommChannel([NotNull] Action<string> action) {
            Requires.NotNull(action, nameof(action));

            ClosePort();
            commPort.DataReceived += SerialDeviceDataReceivedHandler;
            this.action = action;
        }

        private bool IsOpen => commPort.IsOpen;

        public void Dispose() {
            commPort.Close();
            commPort.Dispose();
        }

        internal void ClosePort() {
            if (IsOpen) {
                commPort.Close();
            }
        }

        internal bool OpenPort() {
            bool result = true;
            void HandleException() => result = false;
            try {
                commPort.Open();
            }
            catch (UnauthorizedAccessException) {
                // There are several possible causes.
                // 1. access is denied
                // 2. the current process has already opened it
                // 3. another process has already opened it
                HandleException();
            }
            catch (ArgumentOutOfRangeException) {
                // One or more of the properties for this instance are invalid
                HandleException();
            }
            catch (IOException) {
                // The port is in an invalid state or setting the port state failed.
                HandleException();
            }
            catch (InvalidOperationException) {
                // The port is already open
                HandleException();
            }

            return result;
        }

        internal void WriteData([NotNull] string dataString) {
            Requires.NotNull(dataString, nameof(dataString));

            Console.WriteLine(Resources.GotPrintCommand);
            if (!IsOpen) {
                commPort.Open();
            }

            commPort.Write(dataString);
        }

        private void SerialDeviceDataReceivedHandler([NotNull] object sender, [NotNull] SerialDataReceivedEventArgs e) {
            Requires.NotNull(sender, nameof(sender));

            var sp = (SerialPort)sender;
            action(sp.ReadExisting());
        }
    }
}
