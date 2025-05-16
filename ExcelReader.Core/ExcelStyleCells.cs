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

  public ExcelStyleCells(
    string filePath,
    int headerStart,
    ExcelWorksheet worksheet,
    List<string> columns1,
    List<string> columns2)
  {
    _filePath = filePath;
    _headerStart = headerStart;
    _worksheet = worksheet;
    _columns1 = columns1;
    _columns2 = columns2;
  }

  public void HighlightDifferences(List<CellDifference> differences, string rowKeyColumn)
  {
    using var package = new ExcelPackage(new FileInfo(_filePath));
    _worksheet = package.Workbook.Worksheets[_worksheet.Index];
    // ClearOldHighlights(_worksheet);

    var groupedDiffs = differences
      .GroupBy(cell => cell.RowKey)
      .ToDictionary(group => group.Key, group => group.ToList());

    var keyToExcelRow = new Dictionary<string, int>();
    try
    {
      var keyColumnIndex = GetColumnIndexByName(_worksheet, rowKeyColumn);

      for (var row = _headerStart; row <= _worksheet.Dimension.End.Row; row++)
      {
        var key = _worksheet.Cells[row, keyColumnIndex].Text.Trim();
        if (!string.IsNullOrEmpty(key) && !keyToExcelRow.ContainsKey(key)) keyToExcelRow[key] = row;
      }
    }
    catch (Exception err)
    {
      Console.WriteLine(err);
    }

    // Track how many rows we've inserted after each existing key
    var insertOffsetMap = new Dictionary<string, int>();

    foreach (var (rowKey, cellDiffs) in groupedDiffs)
    {
      var firstDiff = cellDiffs.First();
      int rowIndex;

      if (firstDiff.isNewRow)
      {
          // Get all current keys in order
          // this I will need to understand better (now have bad catch of what is going on under hood)
          var orderedKeys = keyToExcelRow.Keys
            .OrderBy(k => k, Comparer<string>.Create(CompareKeys))
            .ToList();

          // Find the best previous key
          var prevKey = FindPreviousKey(firstDiff.RowKey, orderedKeys);
          if (string.IsNullOrEmpty(prevKey) || !keyToExcelRow.TryGetValue(prevKey, out var prevRowIndex))
            continue;

          // Offset insertion logic
          var offset = insertOffsetMap.TryGetValue(prevKey, out var val) ? val : 0;
          rowIndex = prevRowIndex + 1 + offset;

          _worksheet.InsertRow(rowIndex, 1);

          // Fill values for changed cells
          foreach (var diff in cellDiffs)
          {
            var colIndex = GetColumnIndexByName(_worksheet, diff.ColumnName);
            var cell = _worksheet.Cells[rowIndex, colIndex];
            cell.Value = diff.NewValue;
          }

          // Highlight full row range up to last column from file1
          var howManyColumnToHighlight = 1;
          for (var i = 0; i <= howManyColumnToHighlight; i++)
          {
            var colName = _columns1[i];
            var colIndex = GetColumnIndexByName(_worksheet, colName);
            var cell = _worksheet.Cells[rowIndex, colIndex];

            CellNewValueHighlight(cell, HighlightColorHexLine);
          }

          // Update the offset map and key-to-row map
          insertOffsetMap[prevKey] = offset + 1;
          keyToExcelRow[rowKey] = rowIndex;

          // Shift all subsequent rows in keyToExcelRow
          foreach (var key in keyToExcelRow.Keys.ToList().Where(k => k != rowKey && keyToExcelRow[k] >= rowIndex))
          {
            // if (key != rowKey && keyToExcelRow[key] >= rowIndex)
            // {
              keyToExcelRow[key]++;
            // }
          }
      }

      //TODO logic for finding previous existing key, but I think that works not so good, need different approach.
      // var prevKey = firstDiff.PreviousKey;
      //
      // if (string.IsNullOrEmpty(prevKey) || !keyToExcelRow.TryGetValue(prevKey, out var prevRowIndex))
      //   continue;
      //
      // var offset = insertOffsetMap.TryGetValue(prevKey, out var val) ? val : 0;
      // rowIndex = prevRowIndex + 1 + offset;
      //
      // _worksheet.InsertRow(rowIndex, 1);
      //

      //
      // // track how many items we inserted that goes in row.
      // insertOffsetMap[prevKey] = offset + 1;
      //
      // // Update key cache with the new key at the inserted position
      // keyToExcelRow[rowKey] = rowIndex;
      //
      // // Adjust all later keys’ positions in keyToExcelRow
      // foreach (var key in keyToExcelRow.Keys.ToList())
      //   if (keyToExcelRow[key] >= rowIndex && key != rowKey)
      //     keyToExcelRow[key]++;

      //check if works fine
      
      else if (firstDiff.isWasInFile1)
      {
        rowIndex = GetRowIndexByKey(_worksheet, rowKeyColumn, rowKey);
        if (rowIndex == -1) continue;

        foreach (var diff in cellDiffs)
        {
          var colIndex = GetColumnIndexByName(_worksheet, diff.ColumnName);
          var cell = _worksheet.Cells[rowIndex, colIndex];
          var bg = cell.Style.Fill.BackgroundColor;
          if (bg.Rgb != null)
          {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetAuto();
          }

          cell.Value = diff.NewValue;
          CellNewValueHighlight(cell, HighlightColorHexWasInFirstFile);
        }
      }

      else
      {
        rowIndex = GetRowIndexByKey(_worksheet, rowKeyColumn, rowKey);
        if (rowIndex == -1) continue;

        foreach (var diff in cellDiffs)
        {
          if (diff.isNewRow || diff.isWasInFile1) continue;
          var colIndex = GetColumnIndexByName(_worksheet, diff.ColumnName);
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

  // public void ClearOldHighlights(ExcelWorksheet worksheet)
  // {
  //   var rowsCount = worksheet.Dimension.End.Row;
  //   var columnsCount = worksheet.Dimension.End.Column;
  //   for (int row = _headerStart + 1; row <= rowsCount; row++)
  //   {
  //     for (int col = 1; col <= columnsCount; col++)
  //     {
  //       var cell = worksheet.Cells[row, col];
  //       var bgColor = cell.Style.Fill.BackgroundColor?.Rgb;
  //       if (bgColor == HighlightColorHexCell || bgColor == HighlightColorHexLine || bgColor == HighlightColorHexWasInFirstFile)
  //       {
  //         cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
  //         cell.Style.Fill.BackgroundColor.SetColor(Color.White);
  //       }
  //     }
  //   }
  // }

  private int GetColumnIndexByName(ExcelWorksheet worksheet, string columnName)
  {
    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
      if (worksheet.Cells[_headerStart, col].Text.Trim() == columnName)
        return col;

    // return 0;
    throw new Exception($"column {columnName} not found.");
  }

  private int GetRowIndexByKey(ExcelWorksheet worksheet, string columnName, string keyValue)
  {
    var keyColumnIndex = GetColumnIndexByName(worksheet, columnName);
    for (var row = _headerStart + 1; row <= worksheet.Dimension.End.Row; row++)
      if (worksheet.Cells[row, keyColumnIndex].Text.Trim() == keyValue)
        return row;

    return -1;
  }

  private int CompareKeys(string key1, string key2)
  {
    var parts1 = key1.Split('.');
    var parts2 = key2.Split('.');
    var len = Math.Max(parts1.Length, parts2.Length);

    for (int i = 0; i < len; i++)
    {
      var part1 = i < parts1.Length ? parts1[i] : "0";
      var part2 = i < parts2.Length ? parts2[i] : "0";

      bool isNum1 = int.TryParse(part1, out int num1);
      bool isNum2 = int.TryParse(part2, out int num2);

      if (isNum1 && isNum2)
      {
        if (num1 != num2) return num1.CompareTo(num2);
      }
      else
      {
        int result = string.Compare(part1, part2, StringComparison.OrdinalIgnoreCase);
        if (result != 0) return result;
      }
    }

    return 0;
  }

  private string? FindPreviousKey(string newKey, List<string> existingKeys)
  {
    string? previousKey = null;

    for (int i = 0; i < existingKeys.Count - 1; i++)
    {
      var current = existingKeys[i];
      var next = existingKeys[i + 1];

      // (0030, 0035) < 0 && (0035, 0040) < 0
      if (CompareKeys(current, newKey) < 0 && CompareKeys(newKey, next) < 0)
        return current;
    }

    // If no key in the middle, check if it should go at the end
    if (CompareKeys(newKey, existingKeys.Last()) > 0)
      return existingKeys.Last();

    return previousKey;
  }
}