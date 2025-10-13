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
    /// Логика взаимодействия для StatisticOnePerson.xaml
    /// </summary>
    public partial class StatisticOnePerson : UserControl
    {
        public SeriesCollection SeriesCollection { get; set; }

        public StatisticOnePerson()
        {
            InitializeComponent();
            dpEndDate.SelectedDate = DateTime.Today;
            dpStartDate.SelectedDate = DateTime.Today.AddDays(-30);


            SeriesCollection = new SeriesCollection();
            DataContext = this;
            LoadFoundry();
        }

        private async void LoadFoundry()
        {
            try
            {
                using (var db = new BigFishBDEntities())
                {
                    var foundry = await db.Foundry
                        .OrderBy(f => f.FIO_Foundry)
                        .Select(f => f.FIO_Foundry)
                        .ToListAsync();

                    cbFoundry.ItemsSource = foundry;

                    if (foundry.Any())
                    {
                        cbFoundry.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке литейщиков: {ex.Message}");
            }
        }

        private async void BtnLoadStatistics_Click(object sender, RoutedEventArgs e)
        {
            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите начальную и конечную даты");
                return;
            }

            if (cbFoundry.SelectedItem == null)
            {
                MessageBox.Show("Выберите литейщика");
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
                string selectedFoundry = cbFoundry.SelectedItem.ToString();

                using (var db = new BigFishBDEntities())
                {
                    var foundryData = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.FIO_Foundry == selectedFoundry)
                        .ToListAsync();

                    var totalPacks = foundryData.Sum(dr => dr.Packs2 ?? 0);

                    var totalDefects = foundryData.Sum(dr => dr.FinePacksFoundry ?? 0);

                    double defectPercentage = totalPacks > 0 ? (double)totalDefects / totalPacks * 100 : 0;

                    var salaryData = foundryData
                        .Where(dr => dr.Packs2 > 0 && dr.ArticleFoundry != null)
                        .GroupBy(dr => dr.ArticleFoundry)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0),
                            Price = db.Articles.Where(a => a.Article == g.Key).Select(a => a.PriceFoundry).FirstOrDefault()
                        })
                        .ToList();

                    decimal totalSalary = (decimal)salaryData.Sum(x => x.TotalPacks * (x.Price));

                    decimal fineAmount = 0m;
                    if (totalPacks > 0 && totalDefects > 0)
                    {
                        decimal allowedDefects = (decimal)totalPacks * 0.05m;
                        if ((decimal)totalDefects > allowedDefects)
                        {
                            fineAmount = ((decimal)totalDefects - allowedDefects) * 12m;
                        }
                    }

                    decimal additionalServices = await db.DopFoundry
                        .Where(x => x.FIO_Foundry == selectedFoundry &&
                                   x.DateDop >= startDate && x.DateDop <= endDate)
                        .SumAsync(x => (x.Colvo) * (x.PriceForOne));

                    decimal finalSalary = totalSalary - fineAmount + additionalServices;

                    var bestArticle = foundryData
                        .Where(dr => dr.Packs2 > 0 && dr.ArticleFoundry != null)
                        .GroupBy(dr => dr.ArticleFoundry)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0),
                            TotalDefects = g.Sum(x => x.FinePacksFoundry ?? 0)
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
                        tbFoundryTitle.Text = $"Статистика литейщика: {selectedFoundry}";
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
    }
}
