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

            DateTime startDate = dpStartDate.SelectedDate.Value;
            DateTime endDate = dpEndDate.SelectedDate.Value;
            string foundryName = cbFoundry.SelectedItem.ToString();

            GenerateFoundryReportForOne(startDate, endDate, foundryName);
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
                        .Where(dr => dr.DatePack >= startDate &&
                                    dr.DatePack <= endDate &&
                                    dr.FIO_Foundry == selectedFoundry)
                        .Select(dr => new
                        {
                            Packs2 = dr.Packs2 ?? 0,
                            FinePacksFoundry = dr.FinePacksFoundry ?? 0,
                            ArticleFoundry = dr.ArticleFoundry
                        })
                        .ToListAsync();

                    var totalPacks = foundryData.Sum(dr => dr.Packs2);
                    var totalDefects = foundryData.Sum(dr => dr.FinePacksFoundry);
                    double defectPercentage = totalPacks > 0 ? (double)totalDefects / totalPacks * 100 : 0;

                    decimal totalSalary = 0;
                    var articlesWithPacks = foundryData
                        .Where(dr => dr.Packs2 > 0 && !string.IsNullOrEmpty(dr.ArticleFoundry))
                        .GroupBy(dr => dr.ArticleFoundry)
                        .ToList();

                    foreach (var articleGroup in articlesWithPacks)
                    {
                        var articleName = articleGroup.Key;
                        var totalArticlePacks = articleGroup.Sum(dr => dr.Packs2);

                        var articlePrice = await db.Articles
                            .Where(a => a.Article == articleName)
                            .Select(a => a.PriceFoundry)
                            .FirstOrDefaultAsync();

                        totalSalary += (decimal)totalArticlePacks * articlePrice;
                    }

                    decimal premium = 0m;
                    try
                    {
                        var currentFoundryStandardPacks = await db.DailyReport
                            .Where(dr => dr.DatePack >= startDate &&
                                        dr.DatePack <= endDate &&
                                        dr.FIO_Foundry == selectedFoundry &&
                                        dr.Packs2 > 0 &&
                                        !string.IsNullOrEmpty(dr.ArticleFoundry))
                            .Join(db.Articles,
                                  dr => dr.ArticleFoundry,
                                  a => a.Article,
                                  (dr, a) => new { dr.Packs2, a.Type })
                            .Where(x => x.Type == 1)
                            .SumAsync(x => x.Packs2 ?? 0);

                        var maxStandardPacks = await db.DailyReport
                            .Where(dr => dr.DatePack >= startDate &&
                                        dr.DatePack <= endDate &&
                                        dr.Packs2 > 0 &&
                                        !string.IsNullOrEmpty(dr.ArticleFoundry))
                            .Join(db.Articles,
                                  dr => dr.ArticleFoundry,
                                  a => a.Article,
                                  (dr, a) => new { dr.FIO_Foundry, dr.Packs2, a.Type })
                            .Where(x => x.Type == 1)
                            .GroupBy(x => x.FIO_Foundry)
                            .Select(g => g.Sum(x => x.Packs2 ?? 0))
                            .MaxAsync();

                        if (currentFoundryStandardPacks > 0 &&
                            Math.Abs(currentFoundryStandardPacks - maxStandardPacks) < 0.01)
                        {
                            premium = (decimal)currentFoundryStandardPacks * 3m;
                        }
                    }
                    catch
                    {
                        premium = 0;
                    }

                    decimal fineAmount = 0m;
                    if (totalPacks > 0 && totalDefects > 0)
                    {
                        decimal allowedDefects = (decimal)totalPacks * 0.05m;
                        if ((decimal)totalDefects > allowedDefects)
                        {
                            fineAmount = ((decimal)totalDefects - allowedDefects) * 12m;
                        }
                    }


                    decimal additionalServices = 0;
                    try
                    {
                        var dopServices = await db.DopFoundry
                            .Where(x => x.FIO_Foundry == selectedFoundry &&
                                       x.DateDop >= startDate && x.DateDop <= endDate)
                            .Select(x => new { x.Colvo, x.PriceForOne })
                            .ToListAsync();

                        additionalServices = dopServices.Sum(x => (decimal)(x.Colvo * x.PriceForOne));
                    }
                    catch
                    {
                        additionalServices = 0;
                    }

                    decimal advancePayments = 0;
                    try
                    {
                        var advances = await db.AdvancePayFoundry
                            .Where(x => x.FIO_Foundry == selectedFoundry &&
                                       x.DateAdv >= startDate && x.DateAdv <= endDate)
                            .Select(x => x.AdvancePay)
                            .ToListAsync();

                        advancePayments = advances.Sum(x => (decimal)(x ?? 0));
                    }
                    catch
                    {
                        advancePayments = 0;
                    }

                    decimal finalSalary = totalSalary + premium + additionalServices - fineAmount + advancePayments;

                    var bestArticle = foundryData
                        .Where(dr => dr.Packs2 > 0 && !string.IsNullOrEmpty(dr.ArticleFoundry))
                        .GroupBy(dr => dr.ArticleFoundry)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2),
                            TotalDefects = g.Sum(x => x.FinePacksFoundry)
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
                        tbPremium.Text = premium.ToString("N2") + " ₽";
                        tbFine.Text = fineAmount.ToString("N2") + " ₽";
                        tbAdditionalServices.Text = additionalServices.ToString("N2") + " ₽";
                        tbAdvance.Text = advancePayments.ToString("N2") + " ₽";
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

        private void GenerateFoundryReportForOne(DateTime startDate, DateTime endDate, string foundryName)
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

                    worksheet.Cell(1, 1).Value = $"Отчет по литейщику: {foundryName}";
                    worksheet.Cell(2, 1).Value = $"Период: с {startDate:dd.MM.yyyy} по {endDate:dd.MM.yyyy}";
                    worksheet.Range(1, 1, 2, 1).Style.Font.Bold = true;
                    worksheet.Range(1, 1, 2, 1).Style.Font.FontSize = 14;

                    int currentRow = 4;

                    worksheet.Cell(currentRow, 1).Value = "СДЕЛАННЫЕ АРТИКУЛЫ";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                    currentRow++;

                    var articleHeaders = new[] { "Артикул", "Количество пачек", "Бракованные пачки", "Сумма" };
                    for (int i = 0; i < articleHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = articleHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;

                    var articleData = db.DailyReport
                        .Where(dr => dr.FIO_Foundry == foundryName &&
                                    dr.DatePack >= startDate &&
                                    dr.DatePack <= endDate &&
                                    (dr.Packs2 > 0 || dr.FinePacksFoundry > 0)) 
                        .GroupBy(dr => dr.ArticleFoundry)
                        .Select(g => new
                        {
                            Article = g.Key,
                            TotalPacks = g.Sum(x => x.Packs2 ?? 0),
                            TotalDefects = g.Sum(x => x.FinePacksFoundry ?? 0),
                            Price = db.Articles.Where(a => a.Article == g.Key).Select(a => a.PriceFoundry).FirstOrDefault()
                        })
                        .OrderBy(x => x.Article)
                        .ToList();

                    decimal totalArticleSalary = 0;
                    decimal totalAllPacks = 0;
                    decimal totalAllDefects = 0;

                    foreach (var article in articleData)
                    {
                        decimal articlePrice = article.Price;
                        decimal articleSum = (decimal)article.TotalPacks * articlePrice;
                        totalArticleSalary += articleSum;
                        totalAllPacks += (decimal)article.TotalPacks;
                        totalAllDefects += (decimal)article.TotalDefects;

                        worksheet.Cell(currentRow, 1).Value = article.Article;
                        worksheet.Cell(currentRow, 2).Value = Math.Round((decimal)article.TotalPacks, 2);
                        worksheet.Cell(currentRow, 2).Style.NumberFormat.NumberFormatId = 2;
                        worksheet.Cell(currentRow, 3).Value = Math.Round((decimal)article.TotalDefects, 2);
                        worksheet.Cell(currentRow, 3).Style.NumberFormat.NumberFormatId = 2;
                        worksheet.Cell(currentRow, 4).Value = articleSum;
                        worksheet.Cell(currentRow, 4).Style = moneyStyle;
                        currentRow++;
                    }

                    worksheet.Cell(currentRow, 1).Value = "Итого:";
                    worksheet.Cell(currentRow, 1).Style = summaryStyle;
                    worksheet.Cell(currentRow, 2).Value = Math.Round(totalAllPacks, 2);
                    worksheet.Cell(currentRow, 2).Style = summaryStyle;
                    worksheet.Cell(currentRow, 2).Style.NumberFormat.NumberFormatId = 2;
                    worksheet.Cell(currentRow, 3).Value = Math.Round(totalAllDefects, 2);
                    worksheet.Cell(currentRow, 3).Style = summaryStyle;
                    worksheet.Cell(currentRow, 3).Style.NumberFormat.NumberFormatId = 2;
                    worksheet.Cell(currentRow, 4).Value = totalArticleSalary;
                    worksheet.Cell(currentRow, 4).Style = summaryStyle;
                    worksheet.Cell(currentRow, 4).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;

                    currentRow += 2;

                    worksheet.Cell(currentRow, 1).Value = "РАСЧЕТ ЗАРАБОТНОЙ ПЛАТЫ";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    worksheet.Range(currentRow, 1, currentRow, 3).Merge();
                    currentRow++;

                    decimal fineAmount = 0m;
                    if (totalAllPacks > 0 && totalAllDefects > 0)
                    {
                        decimal allowedDefects = totalAllPacks * 0.05m;
                        if (totalAllDefects > allowedDefects)
                        {
                            fineAmount = (totalAllDefects - allowedDefects) * 12m;
                        }
                    }

                    var additionalServicesData = db.DopFoundry
                        .Where(x => x.FIO_Foundry == foundryName &&
                                   x.DateDop >= startDate && x.DateDop <= endDate)
                        .ToList();

                    decimal additionalServices = 0;
                    foreach (var service in additionalServicesData)
                    {
                        decimal colvo = service.Colvo;
                        decimal price = service.PriceForOne;
                        additionalServices += colvo * price;
                    }

                    var advancePaymentsData = db.AdvancePayFoundry
                        .Where(x => x.FIO_Foundry == foundryName &&
                                   x.DateAdv >= startDate && x.DateAdv <= endDate)
                        .ToList();

                    decimal advancePayments = 0;
                    foreach (var advance in advancePaymentsData)
                    {
                        advancePayments += (decimal)(advance.AdvancePay ?? 0);
                    }

                    decimal premium = 0m;

                    var allFoundryStandardPacks = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && (dr.Packs2 ?? 0) > 0)
                        .GroupBy(dr => dr.FIO_Foundry)
                        .Select(g => new
                        {
                            Foundry = g.Key,
                            StandardPacks = g.Sum(x => db.Articles
                                .Where(a => a.Article == x.ArticleFoundry && a.Type == 1)
                                .Select(a => x.Packs2 ?? 0)
                                .FirstOrDefault())
                        })
                        .ToList();

                    double maxStandardPacks = allFoundryStandardPacks.Max(x => x.StandardPacks);
                    var currentFoundryData = allFoundryStandardPacks.FirstOrDefault(x => x.Foundry == foundryName);
                    double currentFoundryStandardPacks = currentFoundryData?.StandardPacks ?? 0;

                    if (currentFoundryStandardPacks > 0 && Math.Abs(currentFoundryStandardPacks - maxStandardPacks) < 0.01)
                    {
                        premium = (decimal)currentFoundryStandardPacks * 3m;
                    }

                    decimal finalSalary = totalArticleSalary + premium + additionalServices - fineAmount + advancePayments;

                    var salaryItems = new List<SalaryItem>
                {
                    new SalaryItem { Description = "Зарплата за литье:", Value = totalArticleSalary },
                    new SalaryItem { Description = "Премия:", Value = premium },
                    new SalaryItem { Description = "Дополнительные услуги:", Value = additionalServices },
                    new SalaryItem { Description = "Штраф за брак:", Value = -fineAmount },
                    new SalaryItem { Description = "Аванс/удержания:", Value = advancePayments },
                    new SalaryItem { Description = "ИТОГО К ВЫПЛАТЕ:", Value = finalSalary }
                };

                    foreach (var item in salaryItems)
                    {
                        worksheet.Cell(currentRow, 1).Value = item.Description;
                        worksheet.Cell(currentRow, 2).Value = item.Value;
                        worksheet.Cell(currentRow, 2).Style = moneyStyle;

                        if (item.Description == "ИТОГО К ВЫПЛАТЕ:")
                        {
                            worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 1).Style.Font.FontSize = 12;
                            worksheet.Cell(currentRow, 2).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 2).Style.Font.FontSize = 12;
                            worksheet.Cell(currentRow, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                        }

                        currentRow++;
                    }

                    currentRow++;


                    if (additionalServicesData.Any())
                    {
                        worksheet.Cell(currentRow, 1).Value = "ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ";
                        worksheet.Cell(currentRow, 1).Style = headerStyle;
                        worksheet.Range(currentRow, 1, currentRow, 5).Merge();
                        currentRow++;

                        var serviceHeaders = new[] { "Дата", "Услуга", "Количество", "Цена за ед.", "Сумма" };
                        for (int i = 0; i < serviceHeaders.Length; i++)
                        {
                            worksheet.Cell(currentRow, i + 1).Value = serviceHeaders[i];
                            worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                        }
                        currentRow++;

                        foreach (var service in additionalServicesData)
                        {
                            decimal colvo = service.Colvo;
                            decimal price = service.PriceForOne;
                            decimal serviceSum = colvo * price;

                            string serviceDate = service.DateDop.ToString("dd.MM.yyyy") ?? "Не указана";
                            string serviceName = service.Name_Dop ?? "Не указано";

                            worksheet.Cell(currentRow, 1).Value = serviceDate;
                            worksheet.Cell(currentRow, 2).Value = serviceName;
                            worksheet.Cell(currentRow, 3).Value = colvo;
                            worksheet.Cell(currentRow, 4).Value = price;
                            worksheet.Cell(currentRow, 4).Style = moneyStyle;
                            worksheet.Cell(currentRow, 5).Value = serviceSum;
                            worksheet.Cell(currentRow, 5).Style = moneyStyle;
                            currentRow++;
                        }

                        worksheet.Cell(currentRow, 4).Value = "Итого:";
                        worksheet.Cell(currentRow, 4).Style = summaryStyle;
                        worksheet.Cell(currentRow, 5).Value = additionalServices;
                        worksheet.Cell(currentRow, 5).Style = summaryStyle;
                        worksheet.Cell(currentRow, 5).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                        currentRow++;
                    }

                    currentRow++;

                    if (advancePaymentsData.Any())
                    {
                        worksheet.Cell(currentRow, 1).Value = "АВАНСЫ И УДЕРЖАНИЯ";
                        worksheet.Cell(currentRow, 1).Style = headerStyle;
                        worksheet.Range(currentRow, 1, currentRow, 2).Merge();
                        currentRow++;

                        var advanceHeaders = new[] { "Дата", "Сумма аванса" };
                        for (int i = 0; i < advanceHeaders.Length; i++)
                        {
                            worksheet.Cell(currentRow, i + 1).Value = advanceHeaders[i];
                            worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                        }
                        currentRow++;

                        foreach (var advance in advancePaymentsData)
                        {
                            string advanceDate = advance.DateAdv?.ToString("dd.MM.yyyy") ?? "Не указана";
                            decimal advanceAmount = (decimal)(advance.AdvancePay ?? 0);

                            worksheet.Cell(currentRow, 1).Value = advanceDate;
                            worksheet.Cell(currentRow, 2).Value = advanceAmount;
                            worksheet.Cell(currentRow, 2).Style = moneyStyle;
                            currentRow++;
                        }
                        worksheet.Cell(currentRow, 1).Value = "Итого:";
                        worksheet.Cell(currentRow, 1).Style = summaryStyle;
                        worksheet.Cell(currentRow, 2).Value = advancePayments;
                        worksheet.Cell(currentRow, 2).Style = summaryStyle;
                        worksheet.Cell(currentRow, 2).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }

                    // Автоподбор ширины колонок
                    worksheet.Columns().AdjustToContents();

                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_литейщик_{foundryName}_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

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

    public class SalaryItem
    {
        public string Description { get; set; }
        public decimal Value { get; set; }
    }
}
