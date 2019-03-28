using System;
using ClosedXML.Excel;

namespace SimpleXLProfiler
{
    public interface IXLProfilerWriter : IDisposable
    {
        IXLProfilerLog StartProfiling(string description, int column = 1);
    }

    public interface IInternalXLHandler
    {
        IXLWorksheet GetWorksheet();
        int GetRow();
    }
}
