using System;
using System.Globalization;
using System.Threading;
using JetBrains.Annotations;
using Validation;
using static ExcelSerialPortListener.Utilities;
using static System.StringComparison;

namespace ExcelSerialPortListener {
    internal static class Program {
        private static bool gotResponse;

        [NotNull]
        private static string Response { get; set; } = string.Empty;

        [STAThread]
        private static void Main([NotNull] [ItemNotNull] string[] args) {
            if (args.Length != 3) {
                FatalError(Resources.Expected3Arguments);
            }

            var instances = GetExcelInstances();
            if (instances.Count == 0) {
                FatalError(Resources.ExcelIsNotRunningPleaseOpenExcel);
            }

            var cellLocation = new CellLocation(workBookName: args[0], workSheetName: args[1], rangeName: args[2]);

            var ScaleComms = new CommChannel(SetResponse);
            bool commsAreOpen = ScaleComms.OpenPort();
            if (!commsAreOpen) {
                FatalError(Resources.FailedToOpenSerialPortConnection);
            }

            var keyboardListener = new KeyboardListener(() => ScaleComms.WriteData("P\r"));

            var mainThread = new Thread(() => ListenToScale());
            var consoleKeyListener = new Thread(keyboardListener.ListenerKeyBoardEvent);

            consoleKeyListener.Start();
            mainThread.Start();

            while (true) {
                if (gotResponse) {
                    mainThread.Abort();
                    consoleKeyListener.Abort();
                    break;
                }
            }

            var excel = new ExcelComms(cellLocation);

            if (!excel.TryFindWorkbookByName(out var workBook)) {
                FatalError(Resources.ExcelIsNotRunning);
            }

            if (!excel.TryWriteStringToWorksheet(workBook, Response)) {
                FatalError(Resources.FailedToWriteToSpreadsheet);
            }

            ScaleComms.ClosePort();
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

            Response = isTimedOut ? Resources.TimedOut : OnlyDigits(Response);
            gotResponse = true;
        }

        [Pure]
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
            Console.WriteLine(Resources.ReceivedResponse0, Program.Response);
        }
    }
}
