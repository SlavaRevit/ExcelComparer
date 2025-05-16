using ExcelWindow;

class Program
{
  [STAThread] // 👈 This is required for WPF UI
  private static void Main(string[] args)
  {
    var window = new MainWindow();
    window.ShowDialog(); // or Application.Run if needed
  }
}