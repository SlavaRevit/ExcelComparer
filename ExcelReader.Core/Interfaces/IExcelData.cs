namespace ExcelReader.Core.Interfaces;

public interface IExcelData
{
  List<Row> Rows { get; }
  IEnumerable<string> Columns { get; }
  int RowCount { get; }
  int ColumnCount { get; }
  object GetValue(string column, int index);
}