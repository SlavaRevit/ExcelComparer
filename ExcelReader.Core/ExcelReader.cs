using ExcelReader.Core.Interfaces;
using OfficeOpenXml;

namespace ExcelReader.Core;

public class ExcelReader : IExcelReader
{
  private readonly ExcelPackage _excelReader;
  private readonly int _headerIndexStart;


  public ExcelReader(string filePath, int headerIndexStart)
  {
    _headerIndexStart = headerIndexStart;
    _excelReader = new ExcelPackage(new FileInfo(filePath));
  }

  public ExcelData Read(ExcelWorksheet worksheet)
  {
    var workSheet = worksheet;
    workSheet.View.RightToLeft = true;

    var rowCount = workSheet.Dimension.End.Row;
    var columnsCount = workSheet.Dimension.End.Column;

    var columnNames = new List<string>();
    for (var col = 1; col <= columnsCount; col++)
    {
      var headerValue = workSheet.Cells[_headerIndexStart, col].Text?.Trim() ?? $"Column-{col}";
      columnNames.Add(headerValue);
    }

    var resultedRows = new List<Row>();

    for (var row = _headerIndexStart + 1; row <= rowCount; row++)
    {
      var rowData = new Dictionary<string, object>();
      string columnName = null;
      // Get key cell value (e.g., in column 1)
      var keyCell = workSheet.Cells[row, 1];
      var keyValue = keyCell.Text?.Trim();
      
      
      for (var col = 1; col <= columnsCount; col++)
      {
        var cell = workSheet.Cells[row, col];
        columnName = columnNames[col - 1];
        var cellText = cell.Text;
        rowData[columnName] = cellText;
      }

      resultedRows.Add(new Row(rowData));
    }

    return new ExcelData(resultedRows, columnNames);
  }

  public static List<ExcelWorksheet> ReadDataAboutExcelFile(string filePath)
  {
    var reader = new ExcelPackage(new FileInfo(filePath));
    return reader.Workbook.Worksheets.ToList();
  }

  public static List<string> ReadDataAboutColumnInFile(string filePath, ExcelWorksheet workSheetName, int headerIndex = 1)
  {
    var reader = new ExcelPackage(new FileInfo(filePath));
    var worksheet = reader.Workbook.Worksheets.First(sheet => sheet.Name == workSheetName.Name);
    var columnLength = worksheet.Dimension.End.Column;
    var columns = new List<string>();
    for (var i = 1; i <= columnLength; i++)
    {
      var cellValue = worksheet.Cells[headerIndex, i].Text?.Trim();
      if (!string.IsNullOrWhiteSpace(cellValue)) columns.Add(cellValue);
    }

    return columns;
  }
}