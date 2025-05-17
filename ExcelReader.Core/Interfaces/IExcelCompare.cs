namespace ExcelReader.Core.Interfaces;

public interface IExcelCompare
{
  List<CellDifference> Compare(string columnName);
}

public interface IExcelCompareNew
{
  List<CellDifferenceNew> Compare(string columnName);
}
