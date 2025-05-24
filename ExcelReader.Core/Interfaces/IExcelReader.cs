using OfficeOpenXml;

namespace ExcelReader.Core.Interfaces;

internal interface IExcelReader
{
  ExcelData Read(ExcelWorksheet worksheet, int headerIndexStart);
}