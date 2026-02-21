using OfficeOpenXml;

namespace ExcelReader.Core;

public class ExcelHelperMethods
{
  public string NormalizeKey(string key)
  {
    var parts = key.Split('.');
    var normalizedParts = new List<string>();
  
    // Preserve all parts as-is except the last one
    for (var i = 0; i < parts.Length; i++)
    {
      if (i == parts.Length - 1)
      {
        if (int.TryParse(parts[i], out int num))
        {
          if (num != 0)
            normalizedParts.Add(num.ToString());
        }
        else
        {
          normalizedParts.Add(parts[i]);
        }
      }
      else
      {
        normalizedParts.Add(parts[i]);
      }
    }
  
    return string.Join(".", normalizedParts);
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
        var compare = num1.CompareTo(num2);
        if (compare != 0) return compare;
      }
      else
      {
        var compare = string.Compare(part1, part2, StringComparison.OrdinalIgnoreCase);
        if (compare != 0) return compare;
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
  
  public string? FindPreviousKeyNormalized(string newKey, List<string> existingKeys)
  {
    string? previousKey = null;
  
    for (var i = 0; i < existingKeys.Count - 1; i++)
    {
      var current = NormalizeKey(existingKeys[i]);

      if (CompareKeys(current, newKey) == 0)
      {
        return existingKeys[i];
      }
    }

    return previousKey;
  }
  
  
  public string Normalize(object val)
  {
    if (val == null) return string.Empty;
    var str = val.ToString()?
      .Trim()
      .Replace("\u00A0", " ")
      .Replace("\r", "")
      .Replace("\n", "");

    return str == "0" || string.IsNullOrWhiteSpace(str) ? string.Empty : str;
  }

  public bool AreValuesEqual(object v1, object v2)
  {
    return Normalize(v1) == Normalize(v2);
  }

  public int GetColumnIndexByName(ExcelWorksheet worksheet, string columnName, int headerStart)
  {
    var columnsLength = worksheet.Dimension.End.Column;
    for (var col = 1; col <= columnsLength; col++)
    {
      var cell = worksheet.Cells[headerStart, col];
      if (cell.Text.Trim() == columnName)
        return col;
    }

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
  
  public Dictionary<string, int> MapKeysToRowIndices(
    ExcelWorksheet worksheet,
    string keyColumnName,
    int headerStart)
  {
    var keyToExcelRow = new Dictionary<string, int>();
    try
    {
      var keyColumnIndex = GetColumnIndexByName(worksheet, keyColumnName, headerStart);

      var rowLength = worksheet.Dimension.End.Row;
      for (var row = headerStart + 1; row < rowLength; row++)
      {
        var cell = worksheet.Cells[row, keyColumnIndex];
        var key = cell.Text.Trim();

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
}