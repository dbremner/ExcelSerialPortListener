using System;
using System.Globalization;
using System.Threading;
using JetBrains.Annotations;
using Validation;
using static ExcelSerialPortListener.Utilities;

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

            var scaleComms = new CommChannel(SetResponse);
            bool commsAreOpen = scaleComms.OpenPort();
            if (!commsAreOpen) {
                FatalError(Resources.FailedToOpenSerialPortConnection);
            }

            var keyboardListener = new KeyboardListener(() => scaleComms.WriteData("P\r"));
            var scaleListener = new ScaleListener(() => gotResponse = true);

            var mainThread = new Thread(() => scaleListener.ListenToScale());
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

            scaleComms.ClosePort();
        }

        private static void SetResponse([NotNull] string data) {
            Requires.NotNull(data, nameof(data));

            Response = data;
            Console.WriteLine(Resources.ReceivedResponse0, Program.Response);
        }
    }
}
