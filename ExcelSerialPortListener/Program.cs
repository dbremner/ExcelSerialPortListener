using System;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Threading;
using System.Windows.Forms;
using JetBrains.Annotations;

namespace ExcelSerialPortListener {
    internal class Program {
        [NotNull]
        public static string Response { get; set; } = string.Empty;
        private static bool _gotResponse;
        private static CommChannel ScaleComms { get; } = new CommChannel();

        [STAThread]
        private static void Main([ItemNotNull] string[] args) {
            void ErrorMessage(string message) {
                MessageBox.Show(message,
                    nameof(ExcelSerialPortListener), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (args.Length != 3) {
                ErrorMessage("Expected 3 arguments: WorkbookName, WorkSheetName, Range");
                return;
            }

            if (Process.GetProcessesByName("excel").Length == 0) {
                ErrorMessage("Excel is not running, please open Excel with the appropriate spreadsheet.");
                return;
            }

            var (workbookName, worksheetName, rangeName) = (args[0], args[1], args[2]);

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

                var excel = new ExcelComms(workbookName, worksheetName, rangeName);
                excel.TryWriteStringToWorksheet(Response);
            }

            ScaleComms.ClosePort();
        }

        private static void ListenerKeyBoardEvent() {
            Contract.Requires(ScaleComms != null);
            while (true) {
                if (Console.ReadKey(true).Key == ConsoleKey.Spacebar) {
                    Console.WriteLine("Saw pressed key!");
                    ScaleComms.WriteData("P\r");
                }
            }
        }

        private static void ListenToScale(double timeOutInSeconds = 30) {
            ListenToScale(DateTime.Now, timeOutInSeconds);
        }

        private static void ListenToScale(DateTime time, double timeOutInSeconds = 30) {
            Contract.Requires(Response != null);
            var timeOut = time.AddSeconds(timeOutInSeconds);
            var isTimedOut = false;
            do {
                if (Response.Length > 0) {
                    break;
                }

                Thread.Sleep(200);
                isTimedOut = time > timeOut;
            } while (!isTimedOut);

            Response = isTimedOut ? "Timed Out" : OnlyDigits(Response);
            _gotResponse = true;
        }

        private static string OnlyDigits([NotNull] string s) {
            if (s == null) {
                throw new ArgumentNullException(nameof(s));
            }

            Contract.EndContractBlock();
            var onlyDigits = s.Trim();
            var indexOfSpaceG = onlyDigits.IndexOf(" g");
            if (indexOfSpaceG > 0) {
                onlyDigits = onlyDigits.Substring(0, indexOfSpaceG);
            }

            return double.TryParse(onlyDigits, out _) ? onlyDigits : string.Empty;
        }
    }
}
