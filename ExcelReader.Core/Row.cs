namespace ExcelReader.Core;

public class Row
{
  private readonly Dictionary<string, object> _data;

  public Row(Dictionary<string, object> data)
  {
    _data = data;
  }

  public object GetAtColumn(string columnName)
  {
    return _data[columnName];
  }
  
  public string? GetKey(string keyColumn)
  {
    return _data.TryGetValue(keyColumn, out var keyVal) ? keyVal.ToString() : null;
  }
  
}