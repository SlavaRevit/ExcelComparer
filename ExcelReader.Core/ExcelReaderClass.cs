using ExcelReader.Core.Interfaces;
using OfficeOpenXml;

namespace ExcelReader.Core;

public class ExcelReaderClass : IExcelReader
{
  public ExcelData Read(ExcelWorksheet worksheet, int headerIndexStart)
  {
    worksheet.View.RightToLeft = true;

    var rowCount = worksheet.Dimension.End.Row;
    var columnsCount = worksheet.Dimension.End.Column;

    var columnNames = new List<string>();
    for (var col = 1; col <= columnsCount; col++)
    {
      var headerValue = worksheet.Cells[headerIndexStart, col].Text?.Trim() ?? $"Column-{col}";
      columnNames.Add(headerValue);
    }

    var resultedRows = new List<Row>();

    for (var row = headerIndexStart + 1; row <= rowCount; row++)
    {
      var rowData = new Dictionary<string, object>();
      string columnName;
      for (var col = 1; col <= columnsCount; col++)
      {
        columnName = columnNames[col - 1];
        var cellValue = GetCellValueWithFullPrecision(worksheet.Cells[row, col]);
        rowData[columnName] = cellValue;
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
  
  // Helper method to get cell value with full precision
  public static object GetCellValueWithFullPrecision(ExcelRange cell)
  {
    if (cell.Value == null) return null;
    
    if (cell.Value is double doubleValue)
    {
      return doubleValue;
    }
    
    if (cell.Value is decimal decimalValue)
    {
      return decimalValue;
    }
    
    if (cell.Value is DateTime dateValue)
    {
      return dateValue;
    }
    
    return cell.Value;
  }
}