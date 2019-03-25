using System;
using System.Diagnostics;
using ClosedXML.Excel;

namespace SimpleXLProfiler
{
    internal sealed class XLProfilerLog : IDisposable
    {
        private readonly Stopwatch _timeTracker = new Stopwatch();
        private readonly IXLWorksheet _workSheet;
        private readonly string _description;
        private readonly int _row;
        private readonly int _column;

        public XLProfilerLog(string description, XLProfilerWriter profiler, int column)
        {
            _workSheet = profiler.GetWorksheet();
            _row = profiler.GetRow();
            _column = column;
            _description = description;
            _timeTracker.Start();
        }

        private void LogToXL()
        {
            _timeTracker.Stop();
            _workSheet.Cell(_row, _column).SetValue(_description);
            _workSheet.Cell(_row, (_column + 1)).SetValue($"{_timeTracker.ElapsedMilliseconds} ms");
        }

        public void Dispose()
        {
            LogToXL();
        }
    }
}
