using System;

namespace SimpleXLProfiler
{
    public interface IXLProfilerWriter : IDisposable
    {
        IXLProfilerLog StartProfiling(string description, int column = 1);
    }
}
