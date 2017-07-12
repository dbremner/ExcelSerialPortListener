using System;
using System.Globalization;
using System.Threading;
using JetBrains.Annotations;
using Validation;
using static ExcelSerialPortListener.Utilities;

// ReSharper disable HeapView.ClosureAllocation
namespace ExcelSerialPortListener {
    internal static class Program {
        private static int gotResponse;

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

            void GotResponse() {
                _ = Interlocked.Exchange(ref gotResponse, 1);
            }

            IScaleListener scaleListener = new ScaleListener(GotResponse);
            ICommChannel scaleComms = new CommChannel(SetResponse);

            void SetResponse(string data) {
                Requires.NotNull(data, nameof(data));

                scaleListener.Response = data;
                Console.WriteLine(Resources.ReceivedResponse0, scaleListener.Response);
            }

            bool commsAreOpen = scaleComms.OpenPort();
            if (!commsAreOpen) {
                FatalError(Resources.FailedToOpenSerialPortConnection);
            }

            void OnKeyPressed() {
                const string printCommand = "P\r";
                scaleComms.WriteData(printCommand);
            }

            IKeyboardListener keyboardListener = new KeyboardListener(OnKeyPressed);

            void ListenToScale() {
                scaleListener.ListenToScale();
            }

            var mainThread = new Thread(ListenToScale);
            var consoleKeyListener = new Thread(keyboardListener.ListenerKeyBoardEvent);

            consoleKeyListener.Start();
            mainThread.Start();

            while (true) {
                if (gotResponse == 1) {
                    mainThread.Abort();
                    consoleKeyListener.Abort();
                    break;
                }
            }

            IExcelComms excel = new ExcelComms(cellLocation);

            if (!excel.TryFindWorkbookByName(out var workBook)) {
                FatalError(Resources.ExcelIsNotRunning);
            }

            if (!excel.TryWriteStringToWorksheet(workBook, scaleListener.Response)) {
                FatalError(Resources.FailedToWriteToSpreadsheet);
            }

            scaleComms.ClosePort();
        }
    }
}