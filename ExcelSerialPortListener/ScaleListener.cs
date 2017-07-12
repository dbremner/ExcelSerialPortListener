using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using JetBrains.Annotations;
using Validation;
using static System.StringComparison;

namespace ExcelSerialPortListener {
    internal sealed class ScaleListener : IScaleListener {
        [NotNull]
        private readonly Action action;

        public ScaleListener([NotNull] Action action) {
            this.action = action;
        }

        [NotNull]
        public string Response { get; set; } = string.Empty;

        public void ListenToScale(double timeOutInSeconds = 30) {
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
            action();
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
    }
}