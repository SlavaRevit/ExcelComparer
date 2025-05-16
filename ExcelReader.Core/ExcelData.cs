using ExcelReader.Core.Interfaces;
using OfficeOpenXml;

namespace ExcelReader.Core;

public class ExcelData : IExcelData
{
  public List<Row> Rows { get; }
  public IEnumerable<string> Columns { get; }
  public int RowCount => Rows.Count;
  public int ColumnCount => Columns.Count();

  public ExcelData(List<Row> rows, IEnumerable<string> columns)
  {
    Rows = rows;
    Columns = columns;
  }
  
  public object GetValue(string columnName, int index)
  {
    return Rows[index].GetAtColumn(columnName);
  }

  public Dictionary<string, Row> GetRowsByKey(string columnName)
  {
    var result = new Dictionary<string, Row>();

    for (int i = 0; i < Rows.Count; i++)
    {
      var cellValueKey = Rows[i].GetAtColumn(columnName)?.ToString();
      if (!string.IsNullOrEmpty(cellValueKey))
      {
        result[cellValueKey] = Rows[i];
      }
    }

    return result;
  }

  public void AddRow(Row newRow)
  {
    Rows.Add(newRow);
  }
}