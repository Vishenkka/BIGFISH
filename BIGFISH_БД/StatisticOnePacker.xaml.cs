using LiveCharts.Wpf;
using LiveCharts;
using System;
using System.Collections.Generic;
using System.Data.Entity;
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

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для StatisticOnePacker.xaml
    /// </summary>
    public partial class StatisticOnePacker : UserControl
    {
        public SeriesCollection SeriesCollection { get; set; }

        public StatisticOnePacker()
        {
            InitializeComponent();
            dpEndDate.SelectedDate = DateTime.Today;
            dpStartDate.SelectedDate = DateTime.Today.AddDays(-30);


            SeriesCollection = new SeriesCollection();
            DataContext = this;
            LoadPacker();
        }

        private async void LoadPacker()
        {
            try
            {
                using (var db = new BigFishBDEntities())
                {
                    var packer = await db.Packers
                        .OrderBy(f => f.FIO)
                        .Select(f => f.FIO)
                        .ToListAsync();

                    cbPacker.ItemsSource = packer;

                    if (packer.Any())
                    {
                        cbPacker.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке упаковщиц: {ex.Message}");
            }
        }

        private async void BtnLoadStatistics_Click(object sender, RoutedEventArgs e)
        {
            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите начальную и конечную даты");
                return;
            }

            if (cbPacker.SelectedItem == null)
            {
                MessageBox.Show("Выберите упаковщицу!");
                return;
            }

            if (dpStartDate.SelectedDate > dpEndDate.SelectedDate)
            {
                MessageBox.Show("Начальная дата не может быть больше конечной");
                return;
            }

            try
            {

                DateTime startDate = dpStartDate.SelectedDate.Value;
                DateTime endDate = dpEndDate.SelectedDate.Value;
                string selectedPacker = cbPacker.SelectedItem.ToString();

                using (var db = new BigFishBDEntities())
                {
                    var packerData = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.FIO == selectedPacker)
                        .ToListAsync();

                    var totalPacks = packerData.Sum(dr => dr.Packs ?? 0);

                    var totalDefects = packerData.Sum(dr => dr.FinePacks ?? 0);

                    double defectPercentage = totalPacks > 0 ? (double)totalDefects / totalPacks * 100 : 0;

                    var salaryData = packerData
                        .Where(dr => dr.Packs > 0 && dr.ArticlePack != null)
                        .GroupBy(dr => dr.ArticlePack)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs ?? 0),
                            Price = db.Articles.Where(a => a.Article == g.Key).Select(a => a.PricePackers).FirstOrDefault()
                        })
                        .ToList();

                    decimal totalSalary = (decimal)salaryData.Sum(x => x.TotalPacks * (x.Price));


                    var fineAmount = packerData.Sum(dr => dr.FinePacks ?? 0);


                    decimal additionalServices = (decimal)await db.DopPackers
                        .Where(x => x.FIO == selectedPacker &&
                                   x.DateDopPackers >= startDate && x.DateDopPackers <= endDate)
                        .SumAsync(x => (x.Colvo) * (x.PriceForOne));

                    decimal finalSalary = (decimal)totalSalary - (decimal)fineAmount + additionalServices;

                    var bestArticle = packerData
                        .Where(dr => dr.Packs > 0 && dr.ArticlePack != null)
                        .GroupBy(dr => dr.ArticlePack)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs ?? 0),
                            TotalDefects = g.Sum(x => x.FinePacks ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0)
                        .Select(x => new
                        {
                            x.Article,
                            x.TotalPacks,
                            x.TotalDefects,
                            DefectPercentage = (double)x.TotalDefects / x.TotalPacks * 100
                        })
                        .OrderBy(x => x.DefectPercentage)
                        .FirstOrDefault();

                    Dispatcher.Invoke(() =>
                    {
                        tbPackerTitle.Text = $"Статистика упаковщицы: {selectedPacker}";
                        tbTotalPacks.Text = totalPacks.ToString("N0");
                        tbTotalDefects.Text = totalDefects.ToString("N0");
                        tbDefectPercentage.Text = defectPercentage.ToString("F1") + "%";

                        tbSalary.Text = totalSalary.ToString("N2") + " ₽";
                        tbFine.Text = fineAmount.ToString("N2") + " ₽";
                        tbAdditionalServices.Text = additionalServices.ToString("N2") + " ₽";
                        tbFinalSalary.Text = finalSalary.ToString("N2") + " ₽";

                        if (bestArticle != null)
                        {
                            tbBestArticle.Text = bestArticle.Article;
                            tbBestArticleStats.Text = $"{bestArticle.TotalPacks} пачек, {bestArticle.DefectPercentage:F1}% брака";
                        }
                        else
                        {
                            tbBestArticle.Text = "Нет данных";
                            tbBestArticleStats.Text = "";
                        }

                        UpdateChart(totalPacks, totalDefects);
                        tbChartNote.Text = $"Статистика за период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке статистики: {ex.Message}");
            }
        }

        private void UpdateChart(double totalPacks, double totalDefects)
        {
            SeriesCollection.Clear();

            double goodPacks = totalPacks - totalDefects;

            SeriesCollection.Add(new PieSeries
            {
                Title = "Качественные пачки",
                Values = new ChartValues<double> { Math.Round(goodPacks, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y:F0} ({point.Participation:P1})",
                Fill = System.Windows.Media.Brushes.LightGreen,
                Foreground = System.Windows.Media.Brushes.Black
            });

            SeriesCollection.Add(new PieSeries
            {
                Title = "Бракованные пачки",
                Values = new ChartValues<double> { Math.Round(totalDefects, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y:F0} ({point.Participation:P1})",
                Fill = System.Windows.Media.Brushes.LightCoral,
                Foreground = System.Windows.Media.Brushes.Black
            });
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as MainWindow;
            var adminMainWindow = Window.GetWindow(this) as AdminMainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowFirstWindow();
            }
            else
            {
                adminMainWindow.ShowAdminFirstWindow();
            }
        }
    }
}
