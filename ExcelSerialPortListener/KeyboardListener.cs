using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JetBrains.Annotations;
using Validation;

namespace ExcelSerialPortListener {
    internal sealed class KeyboardListener : IKeyboardListener {
        [NotNull]
        private readonly Action action;

        public KeyboardListener([NotNull] Action action) {
            Requires.NotNull(action, nameof(action));
            this.action = action;
        }

        public void ListenerKeyBoardEvent() {
            while (true) {
                if (Console.ReadKey(true).Key == ConsoleKey.Spacebar) {
                    Console.WriteLine(Resources.SawPressedKey);
                    action();
                }
            }
        }
    }
}