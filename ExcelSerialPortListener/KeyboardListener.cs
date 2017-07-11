using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Validation;

namespace ExcelSerialPortListener
{
    internal sealed class KeyboardListener
    {
        public KeyboardListener(CommChannel channel)
        {
            Requires.NotNull(ScaleComms, nameof(ScaleComms));
            ScaleComms = channel;
        }

        private CommChannel ScaleComms { get; }

        internal void ListenerKeyBoardEvent()
        {
            while (true)
            {
                if (Console.ReadKey(true).Key == ConsoleKey.Spacebar)
                {
                    Console.WriteLine(Resources.SawPressedKey);
                    ScaleComms.WriteData("P\r");
                }
            }
        }
    }
}
