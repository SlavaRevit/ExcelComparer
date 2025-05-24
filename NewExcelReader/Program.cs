using ExcelWindow;

namespace NewExcelReader;

public class Program
{
  [STAThread]
  private static void Main(string[] args)
  {
    var window = new MainWindow();
    window.ShowDialog(); 
  }
}