using System;
using System.Configuration;
using System.IO;
using ClosedXML.Excel;

namespace SimpleXLProfiler
{
    public class XLProfilerWriter : IXLProfilerWriter, IInternalXLHandler
    {
        private static readonly Lazy<XLProfilerWriter> lazy = new Lazy<XLProfilerWriter>(() => new XLProfilerWriter());

        public static IXLProfilerWriter Instance { get { return lazy.Value; } }
        private static IInternalXLHandler InternalInstance { get { return lazy.Value; } }

        private string _worksheetTitle;
        private string _filePath;

        private int _row;
        private IXLWorkbook _workBook;
        private IXLWorksheet _workSheet;

        private XLProfilerWriter()
        {
            TryGetConfiguration();
            InitWorkbook(_filePath);
            InitWorksheet();
            _row = _workSheet.LastRowUsed()?.RowNumber() ?? 1;
            CreateHeaderWithStyle();
        }

        public IXLProfilerLog StartProfiling(string description, int column = 1)
        {
            return new XLProfilerLog(description, InternalInstance, column);
        }

        public IXLWorksheet GetWorksheet()
        {
            return _workSheet;
        }

        public int GetRow()
        {
            return ++_row;
        }

        private void InitWorkbook(string filePath)
        {
            _workBook = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook();
        }

        private void InitWorksheet()
        {
            if (_workBook.Worksheets.Contains(_worksheetTitle))
                _workBook.Worksheets.TryGetWorksheet(_worksheetTitle, out _workSheet);
            else
                _workSheet = _workBook.Worksheets.Add(_worksheetTitle);
        }

        private void CreateHeaderWithStyle()
        {
            _row++;
            _workSheet.Row(_row).Style.Fill.BackgroundColor = XLColor.Black;
            _workSheet.Row(_row).Style.Font.FontColor = XLColor.White;
            _workSheet.Row(_row).Style.Font.FontSize = 12;
            _workSheet.Row(_row).Style.Font.Bold = true;

            _workSheet.Cell(_row, 1).SetValue("Description");
            _workSheet.Cell(_row, 2).SetValue("Execution Time");
        }

        public void Dispose()
        {
            _workSheet.Columns().AdjustToContents();
            _workBook.SaveAs(_filePath);
            _workBook.Dispose();
        }

        private void TryGetConfiguration()
        {
            _filePath = ConfigurationManager.AppSettings["SimpleXLProfiler.FilePath"];
            _worksheetTitle = ConfigurationManager.AppSettings["SimpleXLProfiler.WorksheetTitle"];

            if(_filePath == null || _worksheetTitle == null)
                throw new ConfigurationErrorsException("You should provide values in your AppSettings for SimpleXLProfiler.FilePath and SimpleXLProfiler.WorksheetTitle");
        }
    }
}
