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
            CmbUser.ItemsSource = _context.User.ToList(); 
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType)); 

            ChartPayments.ChartAreas.Add(new ChartArea("Main"));

            var currentSeries = new Series("Платежи")
            {
                IsValueShownAsLabel = true
            };
            ChartPayments.Series.Add(currentSeries);

            CmbUser.SelectionChanged += UpdateChart;
            CmbDiagram.SelectionChanged += UpdateChart;

        }
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

                    var paymentSum = _context.Payment
                        .Where(u => u.User == currentUser && u.Category == category)
                        .Sum(u => u.Price * u.Num);

                    currentSeries.Points.AddXY(category.Name, paymentSum);
                }
            }
        }
    }
}
