using ExcelReader.Core.Interfaces;
using OfficeOpenXml;

namespace ExcelReader.Core;

public class ExcelComparerNewVersion : IExcelCompareNew
{
  private readonly ExcelData _originalFile;
  private readonly ExcelData _compareToFile;

  public ExcelComparerNewVersion(ExcelData originalFile, ExcelData compareToFile)
  {
    _originalFile = originalFile;
    _compareToFile = compareToFile;
  }

  public List<CellDifferenceNew> Compare(string columnName)
  {
    var originalFileRows = _originalFile.GetRowsByKey(columnName);
    var compareToFileRows = _compareToFile.GetRowsByKey(columnName);

    var diffs = new List<CellDifferenceNew>();

    foreach (var key in compareToFileRows.Keys)
    {
      var isExistedInOriginalFile = originalFileRows.TryGetValue(key, out var rowOfOriginalFile);

      // new row
      if (!isExistedInOriginalFile)
      {
        var prevKey = FindPreviousKey(key, originalFileRows.Keys.ToList());
        var rowInCompareFile = compareToFileRows[key];

        // Process all columns for this new row
        foreach (var column in _compareToFile.Columns)
        {
          if (!_originalFile.Columns.Contains(column))
            continue;

          var value2 = rowInCompareFile.GetAtColumn(column);
          diffs.Add(new CellDifferenceNew
          {
            RowKey = key,
            ColumnName = column,
            OldValue = null,
            NewValue = value2,
            IsNewRow = true,
            PreviousKey = prevKey,
            IsWasInFile1 = false
          });
        }
      }
      //handle existing row with potential modifications.
      else
      {
        foreach (var column in _compareToFile.Columns)
        {
          if (!_originalFile.Columns.Contains(column))
            continue;

          var value1 = rowOfOriginalFile?.GetAtColumn(column);
          var value2 = compareToFileRows[key].GetAtColumn(column);

          if (!AreValuesEqual(value2, value1))
            diffs.Add(new CellDifferenceNew
            {
              RowKey = key,
              ColumnName = column,
              OldValue = value1,
              NewValue = value2,
              IsNewRow = false,
              PreviousKey = null,
              IsWasInFile1 = false
            });
        }
      }
    }

    // 2. Detect rows that were in file 1 but not in file 2 (deleted rows)
    foreach (var key in originalFileRows.Keys)
    {
      if (!compareToFileRows.ContainsKey(key))
      {
        var row1 = originalFileRows[key];
        foreach (var column in _originalFile.Columns)
        {
          if (!_compareToFile.Columns.Contains(column))
            continue;

          var value1 = row1.GetAtColumn(column);
          diffs.Add(new CellDifferenceNew
          {
            RowKey = key,
            ColumnName = column,
            OldValue = value1,
            NewValue = value1,
            IsNewRow = false,
            PreviousKey = null,
            IsWasInFile1 = true
          });
        }
      }
    }

    return diffs;
  }

  public Dictionary<string, int> MapKeysToRowIndices(
    ExcelWorksheet worksheet,
    string keyColumnName,
    int headerStart)
  {
    var keyToExcelRow = new Dictionary<string, int>();
    try
    {
      var keyColumnIndex = GetColumnIndexByName(worksheet, keyColumnName, headerStart);

      for (var row = headerStart + 1; row <= worksheet.Dimension.End.Row; row++)
      {
        var key = worksheet.Cells[row, keyColumnIndex].Text.Trim();
        if (!string.IsNullOrEmpty(key) && !keyToExcelRow.ContainsKey(key))
          keyToExcelRow[key] = row;
      }
    }
    catch (Exception err)
    {
      Console.WriteLine($"Error mapping keys to row indices: {err.Message}");
    }

    return keyToExcelRow;
  }

  private string Normalize(object val)
  {
    if (val == null) return string.Empty;

    var str = val.ToString().Trim()
      .Replace("\u00A0", " ")
      .Replace("\r", "")
      .Replace("\n", "");

    return str == "0" || string.IsNullOrWhiteSpace(str) ? string.Empty : str;
  }

  private bool AreValuesEqual(object v1, object v2)
  {
    return Normalize(v1) == Normalize(v2);
  }

  public int CompareKeys(string key1, string key2)
  {
    var parts1 = key1.Split('.');
    var parts2 = key2.Split('.');
    var len = Math.Max(parts1.Length, parts2.Length);

    for (var i = 0; i < len; i++)
    {
      var part1 = i < parts1.Length ? parts1[i] : "0";
      var part2 = i < parts2.Length ? parts2[i] : "0";

      var isNum1 = int.TryParse(part1, out var num1);
      var isNum2 = int.TryParse(part2, out var num2);

      if (isNum1 && isNum2)
      {
        if (num1 != num2) return num1.CompareTo(num2);
      }
      else
      {
        var result = string.Compare(part1, part2, StringComparison.OrdinalIgnoreCase);
        if (result != 0) return result;
      }
    }

    return 0;
  }

  public string? FindPreviousKey(string newKey, List<string> existingKeys)
  {
    string? previousKey = null;

    for (var i = 0; i < existingKeys.Count - 1; i++)
    {
      var current = existingKeys[i];
      var next = existingKeys[i + 1];
      
      if (CompareKeys(current, newKey) < 0 && CompareKeys(newKey, next) < 0)
        return current;
    }
    
    
    if (CompareKeys(newKey, existingKeys.Last()) > 0)
      return existingKeys.Last();

    return previousKey;
  }

  public int GetColumnIndexByName(ExcelWorksheet worksheet, string columnName, int headerStart)
  {
    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
      if (worksheet.Cells[headerStart, col].Text.Trim() == columnName)
        return col;
    
    throw new Exception($"column {columnName} not found.");
  }

  public int GetRowIndexByKey(ExcelWorksheet worksheet, string columnName, string keyValue, int headerStart)
  {
    var keyColumnIndex = GetColumnIndexByName(worksheet, columnName, headerStart);
    for (var row = headerStart + 1; row <= worksheet.Dimension.End.Row; row++)
      if (worksheet.Cells[row, keyColumnIndex].Text.Trim() == keyValue)
        return row;

    return -1;
  }
}

public class CellDifferenceNew
{
  public string RowKey { get; set; }
  public string ColumnName { get; set; }
  public object OldValue { get; set; }
  public object NewValue { get; set; }
  public bool IsNewRow { get; set; }
  public string PreviousKey { get; set; }
  public bool IsWasInFile1 { get; set; }
}