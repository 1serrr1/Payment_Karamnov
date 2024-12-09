using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.Integration;


namespace Payment_Karamnov
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_Closing;
            CmbUser.ItemsSource = _context.User.ToList(); // ФИО пользователей
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType)); // Типы диаграммы

            var chart = new Chart();
            chart.ChartAreas.Add(new ChartArea("Main"));

            var currentSeries = new Series("Платежи")
            {
                IsValueShownAsLabel = true
            };
            chart.Series.Add(currentSeries);

            // Добавление диаграммы в WindowsFormsHost
            ChartHost.Child = chart;

            // Привязка событий
            CmbUser.SelectionChanged += UpdateChart;
            CmbDiagram.SelectionChanged += UpdateChart;

            // Сохранение ссылки на диаграмму для дальнейшего использования
            ChartPayments = chart;

        }
        private Chart ChartPayments { get; set; }
        private Karamnov_DB_PaymentEntities2 _context = new Karamnov_DB_PaymentEntities2();

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Вы уверены, что хотите выйти?", "Подтверждение выхода", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.No)
            {
                e.Cancel = true;
            }
        }
        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (CmbUser.SelectedItem is User currentUser && CmbDiagram.SelectedItem is SeriesChartType currentType)
            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault();
                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categoriesList = _context.Category.ToList();
                foreach (var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.Name,
                        _context.Payment.ToList().Where(u => u.User == currentUser
                        && u.Category == category).Sum(u => u.Price * u.Num));
                }
            }
        }


        private void ButtonExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();

                var excelApp = new Excel.Application();
                if (excelApp == null)
                {
                    MessageBox.Show("Excel не установлен на этом компьютере.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                excelApp.SheetsInNewWorkbook = allUsers.Count();
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);

                for (int i = 0; i < allUsers.Count(); i++)
                {

                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[i + 1];
                    worksheet.Name = allUsers[i].FIO;

                    int startRowIndex = 1;

                    worksheet.Cells[1, 1] = "Дата платежа";
                    worksheet.Cells[1, 2] = "Название";
                    worksheet.Cells[1, 3] = "Стоимость";
                    worksheet.Cells[1, 4] = "Количество";
                    worksheet.Cells[1, 5] = "Сумма";

                    Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];
                    columnHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    columnHeaderRange.Font.Bold = true;
                    startRowIndex++;

                    var userCategories = allUsers[i].Payment
                                        .OrderBy(u => u.Date)
                                        .GroupBy(u => u.Category)
                                        .OrderBy(g => g.Key.Name);

                    foreach (var groupCategory in userCategories)
                    {

                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 5]];
                        headerRange.Merge();
                        headerRange.Value = groupCategory.Key.Name;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;

                        foreach (var payment in groupCategory)
                        {
                            worksheet.Cells[startRowIndex, 1] = payment.Date.ToString("dd.MM.yyyy");
                            worksheet.Cells[startRowIndex, 2] = payment.Name;
                            worksheet.Cells[startRowIndex, 3] = payment.Price;
                            ((Excel.Range)worksheet.Cells[startRowIndex, 3]).NumberFormat = "0.00";
                            worksheet.Cells[startRowIndex, 4] = payment.Num;
                            worksheet.Cells[startRowIndex, 5].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                            ((Excel.Range)worksheet.Cells[startRowIndex, 5]).NumberFormat = "0.00";
                            startRowIndex++;
                        }

                        Excel.Range sumLabelRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 4]];
                        sumLabelRange.Merge();
                        sumLabelRange.Value = "ИТОГО:";
                        sumLabelRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                        worksheet.Cells[startRowIndex, 5].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()},E{startRowIndex - 1})";
                        worksheet.Cells[startRowIndex, 5].Font.Bold = true;
                        ((Excel.Range)worksheet.Cells[startRowIndex, 5]).NumberFormat = "0.00";
                        startRowIndex++;

                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRowIndex - 1, 5]];
                        rangeBorders.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        worksheet.Columns.AutoFit();
                    }

                    Excel.Range overallSumLabelRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 4]];
                    overallSumLabelRange.Merge();
                    overallSumLabelRange.Value = "Общий итог:";
                    overallSumLabelRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    overallSumLabelRange.Font.Bold = true;
                    overallSumLabelRange.Font.Color = Excel.XlRgbColor.rgbRed;


                    worksheet.Cells[startRowIndex, 5].Formula = $"=SUM(E2:E{startRowIndex - 1})";
                    worksheet.Cells[startRowIndex, 5].Font.Bold = true;
                    ((Excel.Range)worksheet.Cells[startRowIndex, 5]).NumberFormat = "0.00";
                    worksheet.Cells[startRowIndex, 5].Font.Color = Excel.XlRgbColor.rgbRed;

                    startRowIndex++;

                    Excel.Range overallSumRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRowIndex - 1, 5]];
                    overallSumRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    worksheet.Columns.AutoFit();
                }

                string outputDirectory = @"C:\Temp";
                if (!System.IO.Directory.Exists(outputDirectory))
                {
                    System.IO.Directory.CreateDirectory(outputDirectory);
                }

                string filePath = System.IO.Path.Combine(outputDirectory, "Payments.xlsx");

                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                workbook.SaveAs(filePath);
                workbook.Close();
                excelApp.Quit();

                System.Diagnostics.Process.Start(filePath);

                MessageBox.Show("Экспорт в Excel выполнен успешно!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ButtonExportWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                var allUsers = _context.User.ToList();
                var allCategories = _context.Category.ToList();


                var application = new Word.Application();
                Word.Document document = application.Documents.Add();


                foreach (Word.Section section in document.Sections)
                {
                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    header.Range.Text = $"Дата экспорта: {DateTime.Now:dd.MM.yyyy}";
                    header.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }

                foreach (Word.Section section in document.Sections)
                {
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    footer.Range.Fields.Add(footer.Range, Word.WdFieldType.wdFieldPage);
                    footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }


                foreach (var user in allUsers)
                {

                    Word.Paragraph userParagraph = document.Paragraphs.Add();
                    Word.Range userRange = userParagraph.Range;
                    userRange.Text = user.FIO;
                    userParagraph.set_Style("Заголовок 1");
                    userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    userRange.InsertParagraphAfter();

                    document.Paragraphs.Add();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);

                    paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle =
                        Word.WdLineStyle.wdLineStyleSingle;
                    paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = paymentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Категория";
                    cellRange = paymentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Сумма расходов";
                    paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                    paymentsTable.Rows[1].Range.Font.Size = 14;
                    paymentsTable.Rows[1].Range.Bold = 1;
                    paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    for (int i = 0; i < allCategories.Count(); i++)
                    {
                        var currentCategory = allCategories[i];
                        cellRange = paymentsTable.Cell(i + 2, 1).Range;
                        cellRange.Text = currentCategory.Name;
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;

                        cellRange = paymentsTable.Cell(i + 2, 2).Range;
                        cellRange.Text = user.Payment.ToList()
                            .Where(u => u.Category == currentCategory)
                            .Sum(u => u.Num * u.Price)
                            .ToString("N2") + " руб.";
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;
                    }

                    document.Paragraphs.Add();

                    Payment maxPayment = user.Payment.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                        Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                        maxPaymentRange.Text =
                            $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num):N2} руб. от {maxPayment.Date:dd.MM.yyyy}";
                        maxPaymentParagraph.set_Style("Подзаголовок");
                        maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        maxPaymentRange.InsertParagraphAfter();
                    }

                    document.Paragraphs.Add();


                    Payment minPayment = user.Payment.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                    if (minPayment != null)
                    {
                        Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                        Word.Range minPaymentRange = minPaymentParagraph.Range;
                        minPaymentRange.Text =
                            $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num):N2} руб. от {minPayment.Date:dd.MM.yyyy}";
                        minPaymentParagraph.set_Style("Подзаголовок");
                        minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                        minPaymentRange.InsertParagraphAfter();
                    }

                    if (user != allUsers.LastOrDefault())
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }

                document.SaveAs2(@"C:\Temp\Payments.docx");
                document.SaveAs2(@"C:\Temp\Payments.pdf", Word.WdExportFormat.wdExportFormatPDF);
                application.Visible = true;


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
