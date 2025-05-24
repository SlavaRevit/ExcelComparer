using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReader.Core;

public class ExcelStyleCells
{
  private const string HighlightColorHexCell = "#f2c5e9";
  private const string HighlightColorHexLine = "#95b1ff";
  private const string HighlightColorHexWasInFirstFile = "#d4d4d4";
  private readonly List<string> _columns1;
  private readonly string _filePath;
  private readonly int _headerStart;
  private ExcelWorksheet _worksheet;
  private readonly ExcelHelperMethods _excelHelper;
  private readonly string _selectedColumn;
  private readonly bool? _isChecked;
  private readonly bool? _isMarkCell;

  public ExcelStyleCells(
    string filePath,
    int headerStart,
    ExcelWorksheet worksheet,
    List<string> columns1,
    string selectedColumn,
    bool? isChecked,
    bool? isMarkCell
  )
  {
    _filePath = filePath;
    _headerStart = headerStart;
    _worksheet = worksheet;
    _columns1 = columns1;
    _selectedColumn = selectedColumn;
    _isChecked = isChecked;
    _isMarkCell = isMarkCell;
    _excelHelper = new ExcelHelperMethods();
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
    var keyToExcelRow = _excelHelper.MapKeysToRowIndices(_worksheet, rowKeyColumn, _headerStart);

    foreach (var (rowKey, cellDiffs) in groupedDiffs)
    {
      var firstDiff = cellDiffs.First();
      int rowIndex;

      if (firstDiff.IsNewRow)
      {
        var prevKey = firstDiff.PreviousKey;

        if (string.IsNullOrEmpty(prevKey) || !keyToExcelRow.TryGetValue(prevKey, out var prevRowIndex))
          continue;

        // Offset insertion logic
        var offset = insertOffsetMap.GetValueOrDefault(prevKey, 0);
        rowIndex = prevRowIndex + 1 + offset;

        _worksheet.InsertRow(rowIndex, 1);

        foreach (var diff in cellDiffs)
        {
          var colIndex = _excelHelper.GetColumnIndexByName(_worksheet, diff.ColumnName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          cell.Value = diff.NewValue;
        }

        for (var i = 0; i <= _columns1.Count - 1; i++)
        {
          var colName = _columns1[i];
          var colIndex = _excelHelper.GetColumnIndexByName(_worksheet, colName, _headerStart);
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
        rowIndex = _excelHelper.GetRowIndexByKey(_worksheet, rowKeyColumn, rowKey, _headerStart);
        if (rowIndex == -1) continue;

        foreach (var diff in cellDiffs)
        {
          var colIndex = _excelHelper.GetColumnIndexByName(_worksheet, diff.ColumnName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          cell.Value = diff.OldValue;
        }

        var indexOfSelectedColumn = _excelHelper.GetColumnIndexByName(_worksheet, _selectedColumn, _headerStart);
        var selectedColumnCell = _worksheet.Cells[rowIndex, indexOfSelectedColumn];

        if (_isChecked.HasValue && _isChecked.Value)
        {
          CellNewValueHighlight(selectedColumnCell, HighlightColorHexWasInFirstFile);
        }
      }

      else
      {
        rowIndex = _excelHelper.GetRowIndexByKey(_worksheet, rowKeyColumn, rowKey, _headerStart);
        if (rowIndex == -1) continue;

        foreach (var diff in cellDiffs)
        {
          if (diff.IsNewRow || diff.IsWasInFile1) continue;
          var colIndex = _excelHelper.GetColumnIndexByName(_worksheet, diff.ColumnName, _headerStart);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          cell.Value = diff.NewValue;

          if (_isMarkCell.HasValue && _isMarkCell.Value)
          {
            CellNewValueHighlight(cell, HighlightColorHexCell);
          }

          cell.Style.Numberformat.Format = "0";
        }
      }
    }

    package.Save();
  }

  public void CellNewValueHighlight(ExcelRange cell, string color)
  {
    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
    cell.Style.Fill.BackgroundColor.SetAuto();
    cell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(color));
    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    
    if (double.TryParse(cell.Value?.ToString(), out _))
    {
      cell.Style.Numberformat.Format = "0";
    }
  }
}