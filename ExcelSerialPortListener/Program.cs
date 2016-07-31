using System;
using System.Threading;
using System.Windows.Forms;

namespace ExcelSerialPortListener {
    class Program {
        public static string Response { get; set; } = string.Empty;
        static bool _gotResponse;
        private static CommChannel ScaleComms { get; set; }
        private static bool CommsAreOpen { get; set; }

        static void Main(string[] args) {
            if (args.Length != 3) {
                MessageBox.Show("Expected 3 arguments: WorkbookName, WorkSheetName, Range",
                    nameof(ExcelSerialPortListener), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // args: WorkbookName, WorkSheetName, Range
            string workbookName = args[0];
            string worksheetName = args[1];
            string rangeName = args[2];

            ScaleComms = new CommChannel();

            CommsAreOpen = ScaleComms.OpenPort();
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

        public static void ListenerKeyBoardEvent() {
            do {
                if (Console.ReadKey(true).Key == ConsoleKey.Spacebar) {
                    Console.WriteLine("Saw pressed key!");
                    ScaleComms.WriteData("P\r");
                }
            } while (true);
        }

        public static void ListenToScale(double timeOutInSeconds = 30) {
            var timeOut = DateTime.Now.AddSeconds(timeOutInSeconds);
            var isTimedOut = false;
            do {
                if (Response.Length > 0)
                    break;
                Thread.Sleep(200);
                isTimedOut = DateTime.Now > timeOut;
            } while (!isTimedOut);

            Response = isTimedOut ? "Timed Out" : OnlyDigits(Response);
            _gotResponse = true;
        }

        private static string OnlyDigits(string s) {
            if (s == null) throw new ArgumentNullException(nameof(s));
            var onlyDigits = s.Trim();
            var indexOfSpaceG = onlyDigits.IndexOf(" g");
            if (indexOfSpaceG > 0)
                onlyDigits = onlyDigits.Substring(0, indexOfSpaceG);
            double tester;
            return double.TryParse(onlyDigits, out tester) ? onlyDigits : string.Empty;
        }
    }
}
