namespace ExcelReader.Core.Interfaces;

public interface IExcelCompare
{
  List<CellDifference> Compare(string columnName);
}