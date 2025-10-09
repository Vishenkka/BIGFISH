using System;
using System.Windows;
using System.Windows.Controls;
using System.Data.Entity;
using System.Linq;
using LiveCharts;
using LiveCharts.Wpf;

namespace BIGFISH_БД
{
    public partial class StatisticPage : UserControl
    {
        public SeriesCollection SeriesCollection { get; set; }

        public StatisticPage()
        {
            InitializeComponent();
            dpEndDate.SelectedDate = DateTime.Today;
            dpStartDate.SelectedDate = DateTime.Today.AddDays(-30);
            SeriesCollection = new SeriesCollection(); //для диаграммы 
            DataContext = this;
        }

        private async void BtnLoadStatistics_Click(object sender, RoutedEventArgs e)
        {
            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите начальную и конечную даты");
                return;
            }

            if (dpStartDate.SelectedDate > dpEndDate.SelectedDate)
            {
                MessageBox.Show("Начальная дата не может быть больше конечной");
                return;
            }

            try
            {
                loadingMessage.Visibility = Visibility.Visible;

                DateTime startDate = dpStartDate.SelectedDate.Value;
                DateTime endDate = dpEndDate.SelectedDate.Value;

                using (var db = new BigFishBDEntities())
                {
                    var totalPacks = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .SumAsync(dr => dr.Packs ?? 0 + dr.Packs2 ?? 0);

                    var totalDefects = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .SumAsync(dr =>  dr.FinePacksFoundry ?? 0);

                    double defectPercentage = totalPacks > 0 ? (double)totalDefects / totalPacks * 100 : 0;

                    var bestPacker = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && dr.FIO != null)
                        .GroupBy(dr => dr.FIO)
                        .Select(g => new
                        {
                            PackerName = g.Key,
                            TotalFines = g.Sum(x => x.FinePacks ?? 0),
                            TotalPacks = g.Sum(x => x.Packs ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0) 
                        .OrderBy(x => x.TotalFines)
                        .FirstOrDefaultAsync();

                    var bestFoundry = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && dr.FIO_Foundry != null)
                        .GroupBy(dr => dr.FIO_Foundry)
                        .Select(g => new
                        {
                            FoundryName = g.Key,
                            TotalFines = g.Sum(x => x.FinePacksFoundry ?? 0),
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0)
                        .OrderBy(x => x.TotalFines)
                        .FirstOrDefaultAsync();

                    var worstArticle = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && dr.ArticlePack != null)
                        .GroupBy(dr => dr.ArticlePack)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs ?? 0),
                            TotalFines = g.Sum(x => x.FinePacksFoundry ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0)
                        .Select(x => new
                        {
                            x.Article,
                            x.TotalPacks,
                            x.TotalFines,
                            DefectPercentage = (double)x.TotalFines / x.TotalPacks * 100
                        })
                        .OrderByDescending(x => x.DefectPercentage)
                        .FirstOrDefaultAsync();

                    Dispatcher.Invoke(() =>
                    {
                        tbTotalLures.Text = totalPacks.ToString("N0");
                        tbTotalDefects.Text = totalDefects.ToString("N0");
                        tbDefectPercentage.Text = defectPercentage.ToString("F1") + "%";

                        tbBestPacker.Text = bestPacker?.PackerName ?? "Нет данных";
                        tbBestFoundry.Text = bestFoundry?.FoundryName ?? "Нет данных";

                        if (worstArticle != null)
                        {
                            tbWorstArticle.Text = worstArticle.Article;
                            tbWorstArticlePercentage.Text = $"Процент брака: {worstArticle.DefectPercentage:F1}%";
                        }
                        else
                        {
                            tbWorstArticle.Text = "Нет данных";
                            tbWorstArticlePercentage.Text = "";
                        }

                        UpdateChart((double)totalPacks, totalDefects);

                        tbChartNote.Text = $"Статистика за период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке статистики: {ex.Message}");
            }
            finally
            {
                loadingMessage.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateChart(double totalPacks, double totalDefects)
        {
            SeriesCollection.Clear();

            double goodPacks = totalPacks - totalDefects;

            SeriesCollection.Add(new PieSeries
            {
                Title = "Качественные",
                Values = new ChartValues<double> { Math.Round(goodPacks,2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y} ({point.Participation:P1})",
                Fill = System.Windows.Media.Brushes.LightGreen,
                Foreground = System.Windows.Media.Brushes.Black 
            });

            SeriesCollection.Add(new PieSeries
            {
                Title = "Брак",
                Values = new ChartValues<double> { Math.Round(totalDefects, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y} ({point.Participation:P1})",
                Fill = System.Windows.Media.Brushes.LightCoral,
                Foreground = System.Windows.Media.Brushes.Black 
            });
        }
    }
}