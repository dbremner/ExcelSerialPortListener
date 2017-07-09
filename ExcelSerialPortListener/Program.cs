﻿using System;
using System.Globalization;
using System.Threading;
using JetBrains.Annotations;
using Validation;
using static System.StringComparison;
using static ExcelSerialPortListener.Utilities;

namespace ExcelSerialPortListener {
    internal class Program {
        [NotNull]
        public static string Response { get; set; } = string.Empty;
        private static bool _gotResponse;
        private static CommChannel ScaleComms { get; } = new CommChannel(SetResponse);

        [STAThread]
        private static void Main([ItemNotNull] string[] args) {
            if (args.Length != 3) {
                FatalError("Expected 3 arguments: WorkbookName, WorkSheetName, Range");
            }

            var instances = GetExcelInstances();
            if (instances.Count == 0) {
                FatalError("Excel is not running, please open Excel with the appropriate spreadsheet.");
            }

            var (workbookName, worksheetName, rangeName) = (args[0], args[1], args[2]);

            var cellLocation = new CellLocation(workbookName, worksheetName, rangeName);

            bool CommsAreOpen = ScaleComms.OpenPort();
            if (CommsAreOpen) {
                var mainThread = new Thread(() => ListenToScale());
                var consoleKeyListener = new Thread(ListenerKeyBoardEvent);
                
                consoleKeyListener.Start();
                mainThread.Start();

                while (true) {
                    if (_gotResponse) {
                        mainThread.Abort();
                        consoleKeyListener.Abort();
                        break;
                    } 
                }

                var excel = new ExcelComms(cellLocation);
                excel.TryWriteStringToWorksheet(Response);
            }

            ScaleComms.ClosePort();
        }

        private static void ListenerKeyBoardEvent() {
            Requires.NotNull(ScaleComms, nameof(ScaleComms));

            while (true) {
                if (Console.ReadKey(true).Key == ConsoleKey.Spacebar) {
                    Console.WriteLine("Saw pressed key!");
                    ScaleComms.WriteData("P\r");
                }
            }
        }

        private static void ListenToScale(double timeOutInSeconds = 30) {
            Requires.NotNull(Response, nameof(Response));

            var timeOut = DateTime.Now.AddSeconds(timeOutInSeconds);
            var isTimedOut = false;
            do {
                if (Response.Length > 0) {
                    break;
                }

                Thread.Sleep(200);
                isTimedOut = DateTime.Now > timeOut;
            } while (!isTimedOut);

            Response = isTimedOut ? "Timed Out" : OnlyDigits(Response);
            _gotResponse = true;
        }

        private static string OnlyDigits([NotNull] string s) {
            Requires.NotNull(s, nameof(s));

            var onlyDigits = s.Trim();
            var indexOfSpaceG = onlyDigits.IndexOf(" g", Ordinal);
            if (indexOfSpaceG > 0) {
                onlyDigits = onlyDigits.Substring(0, indexOfSpaceG);
            }

            return double.TryParse(onlyDigits, out _) ? onlyDigits : string.Empty;
        }

        private static void SetResponse([NotNull] string data) {
            Requires.NotNull(data, nameof(data));

            Response = data;
            Console.WriteLine("Received Response: {0}", Program.Response);
        }
    }
}
