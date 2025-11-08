using System;
using System.Windows;
using System.Windows.Controls;
using System.Data.Entity;
using System.Linq;
using LiveCharts;
using LiveCharts.Wpf;
using ClosedXML.Excel;
using System.Diagnostics;

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

            SeriesCollection = new SeriesCollection();
            DataContext = this;
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

        private void BtnGenerateStorageReport_Click(object sender, RoutedEventArgs e)
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

            DateTime startDate = dpStartDate.SelectedDate.Value;
            DateTime endDate = dpEndDate.SelectedDate.Value;

            GenerateStorageReport(startDate, endDate);
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
                        .SumAsync(dr => dr.Packs ?? 0);

                    var totalDefects = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .SumAsync(dr => dr.FinePacksFoundry ?? 0);

                    double defectPercentage = totalPacks > 0 ? (double)totalDefects / totalPacks * 100 : 0;

                    var bestPacker = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && dr.FIO != null)
                        .GroupBy(dr => dr.FIO)
                        .Select(g => new
                        {
                            PackerName = g.Key,
                            TotalPacks = g.Sum(x => x.Packs ?? 0),
                            TotalFines = g.Sum(x => x.FinePacks ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0) 
                        .OrderBy(x => (double)x.TotalFines / x.TotalPacks) 
                        .FirstOrDefaultAsync();

                    var bestFoundry = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && dr.FIO_Foundry != null)
                        .GroupBy(dr => dr.FIO_Foundry)
                        .Select(g => new
                        {
                            FoundryName = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0),
                            TotalFines = g.Sum(x => x.FinePacksFoundry ?? 0)
                        })
                        .Where(x => x.TotalPacks > 0) 
                        .OrderBy(x => (double)x.TotalFines / x.TotalPacks) 
                        .FirstOrDefaultAsync();

                    var worstArticle = await db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && dr.ArticleFoundry != null)
                        .GroupBy(dr => dr.ArticleFoundry)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks2 = g.Sum(x => x.Packs2 ?? 0), 
                            TotalFinePacksFoundry = g.Sum(x => x.FinePacksFoundry ?? 0) 
                        })
                        .Where(x => x.TotalPacks2 > 0) 
                        .Select(x => new
                        {
                            x.Article,
                            TotalAllPacks = x.TotalPacks2 + x.TotalFinePacksFoundry, 
                            TotalDefects = x.TotalFinePacksFoundry, 
                            DefectPercentage = x.TotalPacks2 > 0 ? (double)x.TotalFinePacksFoundry / x.TotalPacks2 * 100 : 0 
                        })
                        .Where(x => x.TotalAllPacks > 0) 
                        .OrderByDescending(x => x.DefectPercentage) 
                        .FirstOrDefaultAsync();

                    Dispatcher.Invoke(() =>
                    {
                        tbTotalLures.Text = totalPacks.ToString("N0");
                        tbTotalDefects.Text = totalDefects.ToString("N0");
                        tbDefectPercentage.Text = defectPercentage.ToString("F1") + "%";

                        if (bestPacker != null)
                        {
                            tbBestPacker.Text = bestPacker.PackerName;
                            tbBestPackerStats.Text = $"{bestPacker.TotalPacks} пачек, {bestPacker.TotalFines} штрафных";
                        }
                        else
                        {
                            tbBestPacker.Text = "Нет данных";
                            tbBestPackerStats.Text = "Недостаточно данных";
                        }

                        if (bestFoundry != null)
                        {
                            tbBestFoundry.Text = bestFoundry.FoundryName;
                            tbBestFoundryStats.Text = $"{bestFoundry.TotalPacks} пачек, {bestFoundry.TotalFines} бракованных";
                        }
                        else
                        {
                            tbBestFoundry.Text = "Нет данных";
                            tbBestFoundryStats.Text = "Недостаточно данных";
                        }
                        if (worstArticle != null)
                        {
                            tbWorstArticle.Text = worstArticle.Article;
                            tbWorstArticlePercentage.Text = $"{worstArticle.TotalAllPacks} всего пачек, {worstArticle.DefectPercentage:F1}% брака";
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
                Values = new ChartValues<double> { Math.Round(goodPacks, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y:F0} ({point.Participation:P1})",
                Fill = System.Windows.Media.Brushes.LightGreen,
                Foreground = System.Windows.Media.Brushes.Black
            });

            SeriesCollection.Add(new PieSeries
            {
                Title = "Брак",
                Values = new ChartValues<double> { Math.Round(totalDefects, 2) },
                DataLabels = true,
                LabelPoint = point => $"{point.Y:F0} ({point.Participation:P1})",
                Fill = System.Windows.Media.Brushes.LightCoral,
                Foreground = System.Windows.Media.Brushes.Black
            });
        }

        private void GenerateStorageReport(DateTime startDate, DateTime endDate)
        {
            try
            {
                using (var db = new BigFishBDEntities())
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Отчет по складам");

                    var headerStyle = workbook.Style;
                    headerStyle.Font.Bold = true;
                    headerStyle.Fill.BackgroundColor = XLColor.Yellow;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽;-#,##0.00 ₽";

                    var summaryStyle = workbook.Style;
                    summaryStyle.Font.Bold = true;
                    summaryStyle.Fill.BackgroundColor = XLColor.Yellow;

                    worksheet.Cell(1, 1).Value = $"Отчет по складам за период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                    worksheet.Range(1, 1, 1, 10).Merge().Style = headerStyle;

                    var articles = db.Articles
                        .OrderBy(a => a.Article)
                        .Select(a => a.Article)
                        .ToList();

                    var storages = db.Storage
                        .OrderBy(s => s.StorageName)
                        .Select(s => s.StorageName)
                        .ToList();

                    int currentRow = 3;

                    worksheet.Cell(currentRow, 1).Value = "Артикул";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;

                    for (int i = 0; i < storages.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 2).Value = storages[i];
                        worksheet.Cell(currentRow, i + 2).Style = headerStyle;
                    }
                    currentRow++;
                    foreach (var article in articles)
                    {
                        worksheet.Cell(currentRow, 1).Value = article;

                        for (int i = 0; i < storages.Count; i++)
                        {
                            var storageName = storages[i];
                            var packsCount = db.DailyReport
                                .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                            dr.ArticlePack == article &&
                                            dr.StorageName == storageName)
                                .Sum(dr => (int?)dr.Packs) ?? 0;

                            worksheet.Cell(currentRow, i + 2).Value = packsCount;
                        }
                        currentRow++;
                    }

                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = "ИТОГО:";
                    worksheet.Cell(currentRow, 1).Style = summaryStyle;

                    for (int i = 0; i < storages.Count; i++)
                    {
                        var storageName = storages[i];
                        var totalStoragePacks = db.DailyReport
                            .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                        dr.StorageName == storageName)
                            .Sum(dr => (int?)dr.Packs) ?? 0;

                        worksheet.Cell(currentRow, i + 2).Value = totalStoragePacks;
                        worksheet.Cell(currentRow, i + 2).Style = summaryStyle;
                    }

                    currentRow += 2;

                    worksheet.Cell(currentRow, 1).Value = "ОБЩАЯ СТАТИСТИКА";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = "Показатель";
                    worksheet.Cell(currentRow, 2).Value = "Количество";
                    worksheet.Cell(currentRow, 3).Value = "Сумма";
                    worksheet.Range(currentRow, 1, currentRow, 3).Style = headerStyle;
                    currentRow++;
                    var totalAllPacks = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .Sum(dr => (int?)dr.Packs) ?? 0;

                    var totalSumData = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .Join(db.Articles,
                              dr => dr.ArticlePack,
                              a => a.Article,
                              (dr, a) => new { dr.Packs, a.PricePackers })
                        .ToList() 
                        .Sum(x => (x.Packs ?? 0) * x.PricePackers);

                    worksheet.Cell(currentRow, 1).Value = "Всего упаковано пачек:";
                    worksheet.Cell(currentRow, 2).Value = totalAllPacks;
                    worksheet.Cell(currentRow, 3).Value = totalSumData;
                    worksheet.Cell(currentRow, 3).Style = moneyStyle;
                    currentRow++;

                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = "ДЕТАЛИЗАЦИЯ ПО АРТИКУЛАМ";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                    currentRow++;

                    worksheet.Cell(currentRow, 1).Value = "Артикул";
                    worksheet.Cell(currentRow, 2).Value = "Количество";
                    worksheet.Cell(currentRow, 3).Value = "Цена";
                    worksheet.Cell(currentRow, 4).Value = "Сумма";
                    worksheet.Range(currentRow, 1, currentRow, 4).Style = headerStyle;
                    currentRow++;

                    foreach (var article in articles)
                    {
                        var articlePacks = db.DailyReport
                            .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate &&
                                        dr.ArticlePack == article)
                            .Sum(dr => (int?)dr.Packs) ?? 0;

                        if (articlePacks > 0)
                        {
                            var articlePrice = db.Articles
                                .Where(a => a.Article == article)
                                .Select(a => a.PricePackers)
                                .FirstOrDefault();

                            var articleSum = articlePacks * articlePrice;

                            worksheet.Cell(currentRow, 1).Value = article;
                            worksheet.Cell(currentRow, 2).Value = articlePacks;
                            worksheet.Cell(currentRow, 3).Value = articlePrice;
                            worksheet.Cell(currentRow, 3).Style = moneyStyle;
                            worksheet.Cell(currentRow, 4).Value = articleSum;
                            worksheet.Cell(currentRow, 4).Style = moneyStyle;
                            currentRow++;
                        }
                    }

                    worksheet.Columns().AdjustToContents();

                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_по_складам_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

                    workbook.SaveAs(tempFilePath);
                    Process.Start(new ProcessStartInfo(tempFilePath) { UseShellExecute = true });

                    MessageBox.Show($"Отчет по складам успешно сформирован!\nФайл: {tempFilePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании отчета по складам: {ex.Message}\n\n{ex.StackTrace}");
            }
        }
    }
}