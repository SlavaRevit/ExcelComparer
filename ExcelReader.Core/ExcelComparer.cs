using ExcelReader.Core.Interfaces;

namespace ExcelReader.Core;

public class ExcelComparer : IExcelCompare
{
  private readonly ExcelData _originalFile;
  private readonly ExcelData _compareToFile;

  public ExcelComparer(ExcelData originalFile, ExcelData compareToFile)
  {
    _originalFile = originalFile;
    _compareToFile = compareToFile;
  }
  
  public List<CellDifference> Compare(string columnName)
  {
    var originalFileRows = _originalFile.GetRowsByKey(columnName);
    var compareToFileRows = _compareToFile.GetRowsByKey(columnName);

    var diffs = new List<CellDifference>();
    // maybe we need to return new dictionary of elements already sorted and compared between each other ?

    foreach (var key in compareToFileRows.Keys)
    {
      var isExistedInOriginalFile = originalFileRows.TryGetValue(key, out var rowOfOriginalFile);
      var compareToFileRow = compareToFileRows[key];
      var insertAfterIndex = -1;

      if (!isExistedInOriginalFile)
      {
        //index of element in file 2 that not present in file 1
        var indexInData2 = _compareToFile.Rows
          .FindIndex(r => r.GetKey(columnName) == key);

        //here I go backward from index that exists in both files
        for (var i = indexInData2 - 1; i >= 0; i--)
        {
          var prevKey = _compareToFile.Rows[i].GetKey(columnName);
          if (prevKey != null && originalFileRows.ContainsKey(prevKey))
          {
            // Find this row's index in _data1
            insertAfterIndex = _originalFile.Rows
              .FindIndex(r => r.GetKey(columnName) == prevKey);
            break;
          }
        }
      }

      _originalFile.AddRow(compareToFileRow);

      foreach (var column in _compareToFile.Columns)
      {
        if (!_originalFile.Columns.Contains(column))
          continue;

        var value1 = isExistedInOriginalFile ? rowOfOriginalFile?.GetAtColumn(column) : null;
        var value2 = compareToFileRows[key].GetAtColumn(column);
        if (!AreValuesEqual(value2, value1))
          diffs.Add(new CellDifference
          {
            RowKey = key,
            ColumnName = column,
            OldValue = value1,
            NewValue = value2,
            isNewRow = !isExistedInOriginalFile,
            PreviousKey = insertAfterIndex != -1 ? _originalFile.Rows[insertAfterIndex].GetKey(columnName) : null
          });
      }
    }

    // 2. Detect rows that were in file 1 but not in file 2 (deleted rows)
    foreach (var key in originalFileRows.Keys)
    {
      // if (key.Length < 3) continue;
      if (!compareToFileRows.ContainsKey(key))
      {
        var row1 = originalFileRows[key];
        foreach (var column in _originalFile.Columns)
        {
          if (!_compareToFile.Columns.Contains(column))
            continue;
    
          var value1 = row1.GetAtColumn(column);
          diffs.Add(new CellDifference
          {
            RowKey = key,
            ColumnName = column,
            OldValue = value1,
            NewValue = value1,
            isNewRow = false,
            isWasInFile1 = true
            // PreviousKey = null // Optional: you could set this to something meaningful if needed
          });
        }
      }
    }

    return diffs;
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
}

public class CellDifference
{
  public string RowKey { get; set; }
  public string ColumnName { get; set; }
  public object OldValue { get; set; }
  public object NewValue { get; set; }
  public bool isNewRow { get; set; }
  public string PreviousKey { get; set; }
  public bool isWasInFile1 { get; set; }
}