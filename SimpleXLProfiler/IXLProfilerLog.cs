using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleXLProfiler
{
    public interface IXLProfilerLog : IDisposable
    {
        void LogToXL();
    }
}
