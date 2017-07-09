﻿using System;
using System.Diagnostics.Contracts;
using System.IO.Ports;
using System.Threading;
using ExcelSerialPortListener.Properties;
using JetBrains.Annotations;

namespace ExcelSerialPortListener {
    public sealed class CommChannel {
        [NotNull] private readonly SerialPort CommPort;
        public bool IsOpen => CommPort.IsOpen;
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
            if(CommPort.IsOpen) CommPort.Close();
            // add listener event handler
            CommPort.DataReceived += SerialDeviceDataReceivedHandler;
        }

        [ContractInvariantMethod]
        private void ObjectInvariant() {
            Contract.Invariant(CommPort != null);
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

        public void WriteData([NotNull] string dataString) {
            if (dataString == null) throw new ArgumentNullException(nameof(dataString));
            Contract.EndContractBlock();
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

        private void SerialDeviceDataReceivedHandler([NotNull] object sender, [NotNull] SerialDataReceivedEventArgs e) {
            if (sender == null) throw new ArgumentNullException(nameof(sender));
            Contract.EndContractBlock();
            var sp = (SerialPort)sender;
            Program.Response = sp.ReadExisting();
            Console.WriteLine("Received Response: {0}", Program.Response);
        }
    }
}
