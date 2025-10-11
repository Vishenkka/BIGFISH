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
    /// Логика взаимодействия для StatisticArticlePage.xaml
    /// </summary>
    public partial class StatisticArticlePage : UserControl
    {
        public SeriesCollection SeriesCollection { get; set; }

        public StatisticArticlePage()
        {
            InitializeComponent();

            dpEndDate.SelectedDate = DateTime.Today;
            dpStartDate.SelectedDate = DateTime.Today.AddDays(-30);

            SeriesCollection = new SeriesCollection();
            DataContext = this;

            LoadArticles();
        }

        private async void LoadArticles()
        {
            try
            {
                using (var db = new BigFishBDEntities())
                {
                    var articles = await db.Articles
                        .OrderBy(a => a.Article)
                        .Select(a => a.Article)
                        .ToListAsync();

                    cbArticles.ItemsSource = articles;

                    if (articles.Any())
                    {
                        cbArticles.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке артикулов: {ex.Message}");
            }
        }

        private async void BtnLoadStatistics_Click(object sender, RoutedEventArgs e)
        {
            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите начальную и конечную даты");
                return;
            }

            if (cbArticles.SelectedItem == null)
            {
                MessageBox.Show("Выберите артикул");
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
                string selectedArticle = cbArticles.SelectedItem.ToString();

                using (var db = new BigFishBDEntities())
                {
                    var totalPacks = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.ArticlePack == selectedArticle)
                        .SumAsync(dr => dr.Packs ?? 0);


                    var totalDefects = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.ArticlePack == selectedArticle)
                        .SumAsync(dr => dr.FinePacksFoundry ?? 0);

                    double totalWithDefects = totalPacks + totalDefects;
                    double defectPercentage = totalWithDefects > 0 ? (double)totalDefects / totalWithDefects * 100 : 0;


                    var bestFoundry = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.ArticleFoundry == selectedArticle &&
                                    dr.FIO_Foundry != null)
                        .GroupBy(dr => dr.FIO_Foundry)
                        .Select(g => new
                        {
                            FoundryName = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0),
                            TotalDefects = g.Sum(x => x.FinePacksFoundry ?? 0) 
                        })
                        .Where(x => x.TotalPacks > 0) 
                        .Select(x => new
                        {
                            x.FoundryName,
                            x.TotalPacks,
                            x.TotalDefects,
                            DefectPercentage = (double)x.TotalDefects / x.TotalPacks * 100 
                        })
                        .OrderBy(x => x.DefectPercentage) 
                        .FirstOrDefaultAsync();


                    Dispatcher.Invoke(() =>
                    {
                        tbArticleTitle.Text = $"Статистика артикула: {selectedArticle}";
                        tbTotalPacks.Text = totalPacks.ToString("N0");
                        tbTotalDefects.Text = totalDefects.ToString("N0");
                        tbDefectPercentage.Text = defectPercentage.ToString("F1") + "%";

                        if (bestFoundry != null)
                        {
                            tbBestFoundry.Text = bestFoundry.FoundryName;
                            tbBestFoundryStats.Text = $"{bestFoundry.TotalPacks} отлито пачек, {bestFoundry.DefectPercentage:F1}% брака";
                        }
                        else
                        {
                            tbBestFoundry.Text = "Нет данных";
                            tbBestFoundryStats.Text = "Нет данных по литейщикам за выбранный период";
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
            SeriesCollection.Add(new PieSeries
            {
                Title = "Качественные пачки",
                Values = new ChartValues<double> { Math.Round(totalPacks, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y:F0} ({point.Participation:P1})",
                Fill = Brushes.LightGreen,
                Foreground = Brushes.Black
            });

            SeriesCollection.Add(new PieSeries
            {
                Title = "Бракованные пачки",
                Values = new ChartValues<double> { Math.Round(totalDefects, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y:F0} ({point.Participation:P1})",
                Fill = Brushes.LightCoral,
                Foreground = Brushes.Black
            });
        }
    }
}
    

