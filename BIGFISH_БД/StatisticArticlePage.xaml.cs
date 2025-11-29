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
using ClosedXML.Excel;
using System.Diagnostics;

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

        private void BtnGenerateReport_Click(object sender, RoutedEventArgs e)
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

            DateTime startDate = dpStartDate.SelectedDate.Value;
            DateTime endDate = dpEndDate.SelectedDate.Value;
            string selectedArticle = cbArticles.SelectedItem.ToString();

            GenerateArticleReport(startDate, endDate, selectedArticle);
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

        private void GenerateArticleReport(DateTime startDate, DateTime endDate, string articleName)
        {
            string connectionString = @"data source=V_ISHENKA\SQLEXPRESS,1433;
                                    initial catalog=BigFishBD;
                                    user id=User1;
                                    password=12345;";

            try
            {
                using (var db = new BigFishBDEntities())
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Отчет");


                    var headerStyle = workbook.Style;
                    headerStyle.Font.Bold = true;
                    headerStyle.Fill.BackgroundColor = XLColor.Yellow;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽;-#,##0.00 ₽";

                    var summaryStyle = workbook.Style;
                    summaryStyle.Font.Bold = true;
                    summaryStyle.Fill.BackgroundColor = XLColor.Yellow;

                    worksheet.Cell(1, 1).Value = $"Отчет по артикулу: {articleName}";
                    worksheet.Cell(2, 1).Value = $"Период: с {startDate:dd.MM.yyyy} по {endDate:dd.MM.yyyy}";
                    worksheet.Range(1, 1, 2, 1).Style.Font.Bold = true;
                    worksheet.Range(1, 1, 2, 1).Style.Font.FontSize = 14;


                    int currentRow = 4;

                    worksheet.Cell(currentRow, 1).Value = "ОБЩАЯ СТАТИСТИКА ПО АРТИКУЛУ";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                    currentRow++;

                    var generalHeaders = new[] { "Показатель", "Значение", "Сумма", "Процент брака" };
                    for (int i = 0; i < generalHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = generalHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;

                    var totalPacks = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.ArticlePack == articleName)
                        .Sum(dr => dr.Packs ?? 0);

                    var totalDefects = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.ArticlePack == articleName)
                        .Sum(dr => dr.FinePacksFoundry ?? 0);

                    var articlePrice = db.Articles
                        .Where(a => a.Article == articleName)
                        .Select(a => a.PricePackers)
                        .FirstOrDefault();

                    decimal totalSum = (decimal)totalPacks * articlePrice;
                    double defectPercentage = totalPacks > 0 ? (double)totalDefects / totalPacks * 100 : 0;


                    worksheet.Cell(currentRow, 1).Value = "Всего пачек";
                    worksheet.Cell(currentRow, 2).Value = totalPacks;
                    worksheet.Cell(currentRow, 3).Value = totalSum;
                    worksheet.Cell(currentRow, 3).Style = moneyStyle;
                    worksheet.Cell(currentRow, 4).Value = defectPercentage.ToString("F1") + "%";
                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = "Бракованные пачки";
                    worksheet.Cell(currentRow, 2).Value = totalDefects;
                    worksheet.Cell(currentRow, 3).Value = "";
                    worksheet.Cell(currentRow, 4).Value = "";
                    currentRow++;

                    currentRow++;


                    worksheet.Cell(currentRow, 1).Value = "СТАТИСТИКА ПО ЛИТЕЙЩИКАМ";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    worksheet.Range(currentRow, 1, currentRow, 5).Merge();
                    currentRow++;


                    var foundryHeaders = new[] { "Литейщик", "Залито пачек", "Бракованных пачек", "Процент брака", "Сумма" };
                    for (int i = 0; i < foundryHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = foundryHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;


                    var foundryData = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                    dr.ArticleFoundry == articleName &&
                                    dr.FIO_Foundry != null)
                        .GroupBy(dr => dr.FIO_Foundry)
                        .Select(g => new
                        {
                            FoundryName = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0),
                            TotalDefects = g.Sum(x => x.FinePacksFoundry ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0)
                        .OrderBy(x => x.FoundryName)
                        .ToList();

                    foreach (var foundry in foundryData)
                    {
                        double foundryDefectPercentage = foundry.TotalPacks > 0 ?
                            (double)foundry.TotalDefects / foundry.TotalPacks * 100 : 0;
                        decimal foundrySum = (decimal)foundry.TotalPacks * articlePrice;

                        worksheet.Cell(currentRow, 1).Value = foundry.FoundryName;
                        worksheet.Cell(currentRow, 2).Value = foundry.TotalPacks;
                        worksheet.Cell(currentRow, 3).Value = foundry.TotalDefects;
                        worksheet.Cell(currentRow, 4).Value = foundryDefectPercentage.ToString("F1") + "%";
                        worksheet.Cell(currentRow, 5).Value = foundrySum;
                        worksheet.Cell(currentRow, 5).Style = moneyStyle;
                        currentRow++;
                    }


                    if (foundryData.Any())
                    {
                        var totalFoundryPacks = foundryData.Sum(x => x.TotalPacks);
                        var totalFoundryDefects = foundryData.Sum(x => x.TotalDefects);
                        decimal totalFoundrySum = (decimal)totalFoundryPacks * articlePrice;
                        double totalFoundryDefectPercentage = totalFoundryPacks > 0 ?
                            (double)totalFoundryDefects / totalFoundryPacks * 100 : 0;

                        worksheet.Cell(currentRow, 1).Value = "Итого по литейщикам:";
                        worksheet.Cell(currentRow, 1).Style = summaryStyle;
                        worksheet.Cell(currentRow, 2).Value = totalFoundryPacks;
                        worksheet.Cell(currentRow, 2).Style = summaryStyle;
                        worksheet.Cell(currentRow, 3).Value = totalFoundryDefects;
                        worksheet.Cell(currentRow, 3).Style = summaryStyle;
                        worksheet.Cell(currentRow, 4).Value = totalFoundryDefectPercentage.ToString("F1") + "%";
                        worksheet.Cell(currentRow, 4).Style = summaryStyle;
                        worksheet.Cell(currentRow, 5).Value = totalFoundrySum;
                        worksheet.Cell(currentRow, 5).Style = summaryStyle;
                        worksheet.Cell(currentRow, 5).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }


                    worksheet.Columns().AdjustToContents();


                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_артикул_{articleName}_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

                    workbook.SaveAs(tempFilePath);
                    Process.Start(new ProcessStartInfo(tempFilePath) { UseShellExecute = true });

                    MessageBox.Show($"Отчет успешно сформирован!\nФайл: {tempFilePath}");
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show($"Ошибка при формировании отчета: {ex.Message}\n\n{ex.StackTrace}");
            }
        }
    }
}
    

