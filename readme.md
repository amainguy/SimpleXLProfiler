# SimpleXLProfiler

## Install SimpleXLProfiler via NuGet

If you want to include SimpleXLProfiler in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/SimpleXLProfiler/)

To install SimpleXLProfiler, run the following command in the Package Manager Console

```
PM> Install-Package SimpleXLProfiler
```

## What is SimpleXLProfiler for?

It is a very lightweight tool that produces an excel file while monitoring 
pieces of code associated with a description.  It's basically a wrapper on `System.Diagnostic.Stopwatch` who write results using [`ClosedXML`](https://github.com/ClosedXML/ClosedXML). If you want to learn further more, take a look at the source code.