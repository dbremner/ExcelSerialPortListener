namespace ExcelSerialPortListener {
    internal interface IScaleListener {
        string Response { get; set; }

        void ListenToScale(double timeOutInSeconds = 30);
    }
}