# SimpleXLProfiler

## What is SimpleXLProfiler for?

It is a very lightweight tool that produces an excel file while monitoring 
pieces of code associated with a description.  It's basically a wrapper on `System.Diagnostic.Stopwatch` who write results using [`ClosedXML`](https://github.com/ClosedXML/ClosedXML). If you want to learn further more, take a look at the [source code](https://github.com/amainguy/SimpleXLProfiler/tree/master/SimpleXLProfiler).

## Install SimpleXLProfiler via NuGet

If you want to include SimpleXLProfiler in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/SimpleXLProfiler/)

To install SimpleXLProfiler, run the following command in the Package Manager Console

```
PM> Install-Package SimpleXLProfiler
```
After installation, just add those two lines to your AppSettings with the values that suit you.

```xml
<add key="SimpleXLProfiler.FilePath" value="C:\profile\profile.xlsx"/>
<add key="SimpleXLProfiler.WorksheetTitle" value="Profiling Results"/>
```
## How to use it

In the method you want to profile, just wrap the code with loggers as this:

```csharp
using (var profiler = XLProfilerWriter.Instance)
{
    var logger = profiler.StartProfiling("do complex things is so long");
    _someService.DoComplexThings();
    _anotherService.DoEvenMoreComplexThings();
    logger.LogToXL();

    logger = profiler.StartProfiling("create documents with style");
    _documentServices.DoMagic();
    logger.LogToXL();        
}
```

and for subroutines, just use the instance with indentation (level) as the second parameter

```csharp
var logger = XLProfilerWriter.Instance.StartProfiling("profile some code", 2)
// some code
logger.LogToXL();
```

the .xlsx document will be created (or updated if it' already there) at disposing time. (ie. the end of the using)
