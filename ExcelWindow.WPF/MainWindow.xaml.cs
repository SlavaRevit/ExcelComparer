using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelReader.Core;
using Microsoft.Win32;
using OfficeOpenXml;

namespace ExcelWindow;

/// <summary>
///   Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow
{
  private string _path1;

  private string _path2;
  private ExcelWorksheet _sheet2;


  public MainWindow()
  {
    InitializeComponent();
    ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
    Console.OutputEncoding = Encoding.UTF8;
  }

  private void BrowseFile1_Click(object sender, RoutedEventArgs e)
  {
    FilePath1Box.Text = SelectExcelFile();
    _path1 = FilePath1Box.Text;
    var worksheets = ExcelReader.Core.ExcelReader.ReadDataAboutExcelFile(_path1);
    WorkSheetFile1.ItemsSource = worksheets;
  }

  private void BrowseFile2_Click(object sender, RoutedEventArgs e)
  {
    FilePath2Box.Text = SelectExcelFile();
    _path2 = FilePath2Box.Text;

    var worksheets = ExcelReader.Core.ExcelReader.ReadDataAboutExcelFile(_path2);
    WorkSheetFile2.ItemsSource = worksheets;
  }

  private string SelectExcelFile()
  {
    var dialog = new OpenFileDialog();
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
    return dialog.ShowDialog() == true ? dialog.FileName : string.Empty;
  }

  private void Compare_Click(object sender, RoutedEventArgs e)
  {
    if (!File.Exists(_path1) || !File.Exists(_path2))
    {
      MessageBox.Show("Please select valid files.");
      return;
    }

    try
    {
      var sheetFile1 = (ExcelWorksheet)WorkSheetFile1.SelectedItem;

      if (sheetFile1 is null)
      {
        MessageBox.Show("Please select sheet's to process.");
        return;
      }

      var selectedColumn = CompareKeyComboBox.SelectedItem as string ?? throw new InvalidOperationException();

      if (string.IsNullOrEmpty(selectedColumn))
      {
        MessageBox.Show("Please select a column to compare.");
        return;
      }

      var headerStart1 = int.TryParse(HeaderFile1Index.Text, out var headerFile1) ? headerFile1 : 1;
      var headerStart2 = int.TryParse(HeaderFile2Index.Text, out var headerFile2) ? headerFile2 : 1;

      var data1 = new ExcelReader.Core.ExcelReader(_path1, headerStart1).Read(sheetFile1);
      var data2 = new ExcelReader.Core.ExcelReader(_path2, headerStart2).Read(_sheet2);

      var comparer = new ExcelComparerNewVersion(data1, data2);
      var diffs = comparer.Compare(selectedColumn);

      // ResultBlock.Text = $"Found {diffs.Count} differences.";

      var columnsData1 = data1.Columns.ToList();
      var columnsData2 = data2.Columns.ToList();

      var styler = new ExcelStyleCells(_path1, headerFile1, sheetFile1, columnsData1, columnsData2, comparer, selectedColumn);
      styler.HighlightDifferences(diffs, selectedColumn);

      MessageBox.Show("Comparison done. File updated.");
    }
    catch (Exception ex)
    {
      MessageBox.Show($"Error: {ex.Message}");
    }
    finally
    {
      Close();
    }
  }

  private void WorkSheetFile2_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
  {
    var selectedSheet = (ExcelWorksheet)WorkSheetFile2.SelectedItem;
    if (selectedSheet == null) return;
    _sheet2 = selectedSheet;
    var headerStart2 = int.TryParse(HeaderFile2Index.Text, out var headerFile2) ? headerFile2 : 1;
    var columns = ExcelReader.Core.ExcelReader
      .ReadDataAboutColumnInFile(_path2, selectedSheet, headerStart2);
    CompareKeyComboBox.ItemsSource = columns;
  }
}