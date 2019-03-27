using System;

namespace SimpleXLProfiler
{
    public interface IXLProfilerLog : IDisposable
    {
        void LogToXL();
    }
}
