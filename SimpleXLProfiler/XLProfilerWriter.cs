using System;
using System.Configuration;
using System.IO;
using ClosedXML.Excel;

namespace SimpleXLProfiler
{
    public class XLProfilerWriter : IDisposable
    {
        private static readonly Lazy<XLProfilerWriter> lazy = new Lazy<XLProfilerWriter>(() => new XLProfilerWriter());

        public static XLProfilerWriter Instance { get { return lazy.Value; } }

        private const string WorksheetTitle = "Profiling results";
        private const string FilePath = "C:\\profile\\profileResult.xlsx";

        private int _row;
        private XLWorkbook _workBook;
        private IXLWorksheet _workSheet;
        private XLProfilerLog _profilerLog;

        private XLProfilerWriter()
        {
            InitWorkbook(FilePath);
            InitWorksheet();
            _row = _workSheet.LastRowUsed()?.RowNumber() ?? 1;
            CreateHeaderWithStyle();
        }

        public void StartProfiling(string description, int column = 1)
        {
            _profilerLog = new XLProfilerLog(description, this, column);
        }

        public void StopProfiling()
        {
            _profilerLog.Dispose();
        }

        internal IXLWorksheet GetWorksheet()
        {
            return _workSheet;
        }

        internal int GetRow()
        {
            return ++_row;
        }

        private void InitWorkbook(string filePath)
        {
            _workBook = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook();
        }

        private void InitWorksheet()
        {
            if (_workBook.Worksheets.Contains(WorksheetTitle))
                _workBook.Worksheets.TryGetWorksheet(WorksheetTitle, out _workSheet);
            else
                _workSheet = _workBook.Worksheets.Add(WorksheetTitle);
        }

        private void CreateHeaderWithStyle()
        {
            _row++;
            _workSheet.Row(_row).Style.Fill.BackgroundColor = XLColor.Black;
            _workSheet.Row(_row).Style.Font.FontColor = XLColor.White;
            _workSheet.Row(_row).Style.Font.FontSize = 12;
            _workSheet.Row(_row).Style.Font.Bold = true;

            _workSheet.Cell(_row, 1).SetValue("Description");
            _workSheet.Cell(_row, 2).SetValue("Temps d'exécution");
        }

        public void Dispose()
        {
            _workSheet.Columns().AdjustToContents();
            _workBook.SaveAs(FilePath);
            _workBook.Dispose();
        }
    }
}
