using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PeerReviewCombinator
{
    public class ExcelWriter : IDisposable
    {
        private readonly string _outputFile;
        private readonly IWorkbook _workbook;
        private readonly ISheet _sheet;

        private int _currentRowIndex;

        public ExcelWriter(string outputFile, List<string> colNames)
        {
            _outputFile = outputFile;

            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet("Default Sheet");

            var row = _sheet.CreateRow(0);
            for (var colIndex = 0; colIndex < colNames.Count; colIndex++)
            {
                row.CreateCell(colIndex).SetCellValue(colNames[colIndex]);
            }

            _currentRowIndex = 1;
        }

        public void WriteRow(List<string> values)
        {
            var row = _sheet.CreateRow(_currentRowIndex++);
            for (int colIndex = 0; colIndex < values.Count; colIndex++)
                row.CreateCell(colIndex).SetCellValue(values[colIndex]);
        }

        public void Finish()
        {
            FileStream fileStream = File.Create(_outputFile);
            _workbook.Write(fileStream, false);
            fileStream.Close();
        }

        void IDisposable.Dispose() => _workbook.Dispose();
    }
}
