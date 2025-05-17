using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReader.Core;

public class ExcelStyleCells
{
  private const string HighlightColorHexCell = "#f2c5e9";
  private const string HighlightColorHexLine = "#7091ff";
  private const string HighlightColorHexWasInFirstFile = "#a3a3a3";
  private readonly List<string> _columns1;
  private readonly List<string> _columns2;
  private readonly string _filePath;
  private readonly int _headerStart;
  private ExcelWorksheet _worksheet;
  private readonly ExcelComparerNewVersion _comparer;


  public ExcelStyleCells(
    string filePath,
    int headerStart,
    ExcelWorksheet worksheet,
    List<string> columns1,
    List<string> columns2,
    ExcelComparerNewVersion comparer
  )
  {
    _filePath = filePath;
    _headerStart = headerStart;
    _worksheet = worksheet;
    _columns1 = columns1;
    _columns2 = columns2;
    _comparer = comparer;
  }

  public void HighlightDifferences(List<CellDifferenceNew> differences, string rowKeyColumn)
  {
    using var package = new ExcelPackage(new FileInfo(_filePath));
    _worksheet = package.Workbook.Worksheets[_worksheet.Index];

    var groupedDiffs = differences
      .GroupBy(cell => cell.RowKey)
      .ToDictionary(group
        => group.Key, group => group.ToList());

    // Track how many rows we've inserted after each existing key
    var insertOffsetMap = new Dictionary<string, int>();
    var keyToExcelRow = _comparer.MapKeysToRowIndices(_worksheet, rowKeyColumn, _headerStart);

    foreach (var (rowKey, cellDiffs) in groupedDiffs)
    {
      var firstDiff = cellDiffs.First();
      int rowIndex;

      if (firstDiff.IsNewRow)
      {
        // Use the comparer to find the previous key
        var prevKey = firstDiff.PreviousKey;
        if (string.IsNullOrEmpty(prevKey) || !keyToExcelRow.TryGetValue(prevKey, out var prevRowIndex))
          continue;

        // Offset insertion logic
        var offset = insertOffsetMap.GetValueOrDefault(prevKey, 0);
        rowIndex = prevRowIndex + 1 + offset;

        _worksheet.InsertRow(rowIndex, 1);
        
        // Fill values for changed cells
        foreach (var diff in cellDiffs)
        {
          var colIndex = _comparer.GetColumnIndexByName(_worksheet, diff.ColumnName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          cell.Value = diff.NewValue;
        }
        
        for (var i = 0; i <= _columns1.Count - 1; i++)
        {
          var colName = _columns1[i];
          var colIndex = _comparer.GetColumnIndexByName(_worksheet, colName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];

          CellNewValueHighlight(cell, HighlightColorHexLine);
        }

        // Update the offset map and key-to-row map
        insertOffsetMap[prevKey] = offset + 1;
        keyToExcelRow[rowKey] = rowIndex;

        // Shift all subsequent rows in keyToExcelRow
        foreach (var key in keyToExcelRow.Keys
                   .ToList()
                   .Where(k => k != rowKey && keyToExcelRow[k] >= rowIndex))
        {
          keyToExcelRow[key]++;
        }
      }

      else if (firstDiff.IsWasInFile1)
      {
        rowIndex = _comparer.GetRowIndexByKey(_worksheet, rowKeyColumn, rowKey, _headerStart);
        if (rowIndex == -1) continue;

        foreach (var diff in cellDiffs)
        {
          var colIndex = _comparer.GetColumnIndexByName(_worksheet, diff.ColumnName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          cell.Value = diff.NewValue;
        }

        for (var i = 0; i <= _columns1.Count - 1; i++)
        {
          var colName = _columns1[i];
          var colIndex = _comparer.GetColumnIndexByName(_worksheet, colName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          
          var bg = cell.Style.Fill.BackgroundColor;
          if (bg.Rgb != null)
          {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetAuto();
          }

          CellNewValueHighlight(cell, HighlightColorHexWasInFirstFile);
        }
      }

      else
      {
        rowIndex = _comparer.GetRowIndexByKey(_worksheet, rowKeyColumn, rowKey, _headerStart);
        if (rowIndex == -1) continue;

        foreach (var diff in cellDiffs)
        {
          if (diff.IsNewRow || diff.IsWasInFile1) continue;
          var colIndex = _comparer.GetColumnIndexByName(_worksheet, diff.ColumnName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          cell.Value = diff.NewValue;
          CellNewValueHighlight(cell, HighlightColorHexCell);
        }
      }
    }

    package.Save();
  }

  public void CellNewValueHighlight(ExcelRange cell, string color)
  {
    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
    cell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(color));
    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
  }
}