using ExcelReader.Core.Interfaces;

namespace ExcelReader.Core;

public class ExcelComparerNewVersion : IExcelCompareNew
{
  private readonly ExcelData _originalFile;
  private readonly ExcelData _compareToFile;
  private readonly ExcelHelperMethods _excelHelper;


  public ExcelComparerNewVersion(
    ExcelData originalFile,
    ExcelData compareToFile
    )
  {
    _originalFile = originalFile;
    _compareToFile = compareToFile;
    _excelHelper = new ExcelHelperMethods();
  }

  public List<CellDifferenceNew> Compare(string columnName)
  {
    var originalRaw = _originalFile.GetRowsByKey(columnName);
    var compareRaw = _compareToFile.GetRowsByKey(columnName);

    var diffs = new List<CellDifferenceNew>();

    foreach (var key in compareRaw.Keys)
    {
      var isExistedInOriginalFile = originalRaw
        .TryGetValue(key, out var rowOfOriginalFile);
      
      var prevKey = _excelHelper
        .FindPreviousKey(key, originalRaw.Keys.ToList());
      
      if (!isExistedInOriginalFile && prevKey is not null)
      {
        foreach (var column in _compareToFile.Columns)
        {
          if (!_originalFile.Columns.Contains(column))
            continue;

          var value2 = compareRaw[key].GetAtColumn(column);
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

      if (!isExistedInOriginalFile && prevKey is null)
      {
        var rowInCompareFile = compareRaw[key];
        var newKey = _excelHelper
          .FindPreviousKeyNormalized(key, originalRaw.Keys.ToList());
        
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
            PreviousKey = newKey,
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

          if (rowOfOriginalFile is null) continue;
          
          var value1 = rowOfOriginalFile.GetAtColumn(column);
          var value2 = compareRaw[key].GetAtColumn(column);
          
          if (!_excelHelper.AreValuesEqual(value2, value1))
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
    foreach (var key in originalRaw.Keys)
    {
      if (!compareRaw.ContainsKey(key))
      {
        foreach (var column in _originalFile.Columns)
        {
          if (!_compareToFile.Columns.Contains(column))
            continue;

          var value1 = originalRaw[key].GetAtColumn(column);
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
    // }

    return diffs;
  }
}

public class CellDifferenceNew
{
  public string? RowKey { get; set; }
  public string? ColumnName { get; set; }
  public object? OldValue { get; set; }
  public object? NewValue { get; set; }
  public bool IsNewRow { get; set; }
  public string? PreviousKey { get; set; }
  public bool IsWasInFile1 { get; set; }
}