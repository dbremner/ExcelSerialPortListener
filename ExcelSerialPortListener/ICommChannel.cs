using System;
using JetBrains.Annotations;

namespace ExcelSerialPortListener {
    internal interface ICommChannel : IDisposable {
        void ClosePort();

        bool OpenPort();

        void WriteData([NotNull] string dataString);
    }
}