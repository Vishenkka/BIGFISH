using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
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
using System.Windows.Shapes;
using ClosedXML.Excel;
using Dapper;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.IO;


namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для DatePickPackers.xaml
    /// </summary>
    public partial class DatePickPackers : Window
    {
        public DatePickPackers()
        {
            InitializeComponent();
        }

        private void GeneratePackerReport(DateTime startDate, DateTime endDate)
        {
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

                    var orangeStyle = workbook.Style;
                    orangeStyle.Fill.BackgroundColor = XLColor.Yellow;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽;-#,##0.00 ₽";

                    //ТАБЛИЦА 1
                    var articleDetails = db.Packers
                        .Select(p => new
                        {
                            Упаковщица = p.FIO,
                            Данные = db.DailyReport
                                .Where(dr => dr.FIO == p.FIO &&
                                            dr.DatePack >= startDate &&
                                            dr.DatePack <= endDate)
                                .GroupBy(dr => dr.ArticlePack)
                                .Select(g => new
                                {
                                    Article = g.Key,
                                    Пачки = g.Sum(x => x.Packs ?? 0),
                                    Штрафы = g.Sum(x => x.FinePacks ?? 0)
                                })
                                .ToList()
                        })
                        .ToList();

                    var allArticles = db.Articles
                        .Select(a => a.Article)
                        .Distinct()
                        .OrderBy(a => a)
                        .ToList();

                    var articlePrices = db.Articles
                        .ToDictionary(a => a.Article, a => a.PricePackers);


                    int currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Упаковщица";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;

                    for (int i = 0; i < allArticles.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 2).Value = allArticles[i];
                        worksheet.Cell(currentRow, i + 2).Style = headerStyle;
                    }


                    currentRow++;
                    worksheet.Row(currentRow).Style = orangeStyle;
                    for (int i = 0; i < allArticles.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 2).Value = articlePrices[allArticles[i]];
                        worksheet.Cell(currentRow, i + 2).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }

                    currentRow++;
                    foreach (var detail in articleDetails)
                    {
                        worksheet.Cell(currentRow, 1).Value = detail.Упаковщица;

                        var articleDict = detail.Данные
                            .GroupBy(d => d.Article)
                            .ToDictionary(
                                g => g.Key,
                                g => Math.Round(g.Sum(x => x.Пачки), 2)
                            );

                        for (int i = 0; i < allArticles.Count; i++)
                        {
                            var article = allArticles[i];
                            worksheet.Cell(currentRow, i + 2).Value =
                                articleDict.ContainsKey(article) ? articleDict[article] : 0;
                        }
                        currentRow++;
                    }


                    worksheet.Cell(currentRow, 1).Value = "Итого:";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    for (int col = 2; col <= allArticles.Count + 1; col++)
                    {
                        worksheet.Cell(currentRow, col).FormulaA1 = $"ROUND(SUM({worksheet.Cell(3, col).Address}:{worksheet.Cell(currentRow - 1, col).Address}), 2)";
                        worksheet.Cell(currentRow, col).Style = headerStyle;
                        worksheet.Cell(currentRow, col).Style.NumberFormat.NumberFormatId = 2;
                    }

                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = "Итого сумма:";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    for (int col = 2; col <= allArticles.Count + 1; col++)
                    {
                        var sumAddress = worksheet.Cell(currentRow - 1, col).Address;
                        var priceAddress = worksheet.Cell(2, col).Address;
                        worksheet.Cell(currentRow, col).FormulaA1 = $"{sumAddress} * {priceAddress}";
                        worksheet.Cell(currentRow, col).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }
                    currentRow += 2;

                    //ТАБЛИЦА 2
                    var allPackers = db.Packers
                        .Where(p =>
                            db.DailyReport.Any(dr => dr.FIO == p.FIO && dr.DatePack >= startDate && dr.DatePack <= endDate) ||
                            db.AdvancePayPackers.Any(ap => ap.FIO == p.FIO && ap.DateAdv >= startDate && ap.DateAdv <= endDate) ||
                            db.DopPackers.Any(dp => dp.FIO == p.FIO && dp.DateDopPackers >= startDate && dp.DateDopPackers <= endDate)
                        )
                        .Select(p => p.FIO)
                        .Distinct()
                        .ToList();

                    var dailyReportData = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate && allPackers.Contains(dr.FIO))
                        .GroupBy(dr => dr.FIO)
                        .Select(g => new
                        {
                            FIO = g.Key,
                            ВсегоПачек = g.Sum(x => x.Packs ?? 0),
                            ВсегоШтрафов = g.Sum(x => x.FinePacks ?? 0),
                            Премия = g.Sum(x => x.FinePacksFoundry ?? 0) * 2,
                            Упаковка = g.Sum(x => (x.Packs ?? 0) * (x.Articles.PricePackers))
                        })
                        .ToList();

                    var dopUslugiData = db.DopPackers
                        .Where(dp => dp.DateDopPackers >= startDate && dp.DateDopPackers <= endDate && allPackers.Contains(dp.FIO))
                        .GroupBy(dp => dp.FIO)
                        .Select(g => new
                        {
                            FIO = g.Key,
                            DopSum = g.Sum(x => (x.Colvo ?? 0) * (x.PriceForOne ?? 0))
                        })
                        .ToList();

                    var advancePaymentsData = db.AdvancePayPackers
                        .Where(ap => ap.DateAdv >= startDate && ap.DateAdv <= endDate && allPackers.Contains(ap.FIO))
                        .GroupBy(ap => ap.FIO)
                        .Select(g => new
                        {
                            FIO = g.Key,
                            TotalAdvance = g.Sum(x => x.AdvancePay ?? 0)
                        })
                        .ToList();

                    var reports = allPackers
                        .Select(fio => new
                        {
                            Упаковщица = fio,
                            DailyData = dailyReportData.FirstOrDefault(d => d.FIO == fio),
                            DopData = dopUslugiData.FirstOrDefault(d => d.FIO == fio),
                            AdvanceData = advancePaymentsData.FirstOrDefault(d => d.FIO == fio)
                        })
                        .Select(x => new
                        {
                            Упаковщица = x.Упаковщица,
                            Пачки = Math.Round(x.DailyData?.ВсегоПачек ?? 0, 2),
                            Штрафы = Math.Round((x.DailyData?.ВсегоШтрафов ?? 0) * 50, 2),
                            Премия = x.DailyData?.Премия ?? 0,
                            Упаковка = x.DailyData?.Упаковка ?? 0,
                            Допы = x.DopData?.DopSum ?? 0,
                            АвансУдержания = x.AdvanceData?.TotalAdvance ?? 0
                        })
                        .OrderBy(x => x.Упаковщица)
                        .ToList();

                    var summaryHeaders = new[] { "Упаковщица", "Упаковка", "Количество пачек", "Штрафы", "Премия", "Допы", "Аванс-удержания", "Итого" };
                    for (int i = 0; i < summaryHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = summaryHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;

                    foreach (var report in reports)
                    {
                        worksheet.Cell(currentRow, 1).Value = report.Упаковщица;
                        worksheet.Cell(currentRow, 2).Value = report.Упаковка;
                        worksheet.Cell(currentRow, 2).Style = moneyStyle;
                        worksheet.Cell(currentRow, 3).Value = report.Пачки;
                        worksheet.Cell(currentRow, 3).Style.NumberFormat.NumberFormatId = 2;
                        worksheet.Cell(currentRow, 4).Value = report.Штрафы;
                        worksheet.Cell(currentRow, 4).Style = moneyStyle;
                        worksheet.Cell(currentRow, 5).Value = report.Премия;
                        worksheet.Cell(currentRow, 5).Style = moneyStyle;
                        worksheet.Cell(currentRow, 6).Value = report.Допы;
                        worksheet.Cell(currentRow, 6).Style = moneyStyle;
                        worksheet.Cell(currentRow, 7).Value = report.АвансУдержания;
                        worksheet.Cell(currentRow, 7).Style = moneyStyle;
                        worksheet.Cell(currentRow, 8).Value = report.Упаковка - report.Штрафы + report.Премия + report.Допы + report.АвансУдержания;
                        worksheet.Cell(currentRow, 8).Style = moneyStyle;
                        currentRow++;
                    }

                    worksheet.Cell(currentRow, 1).Value = "Итого:";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;

                    for (int col = 2; col <= 8; col++)
                    {
                        var addressStart = worksheet.Cell(currentRow - reports.Count, col).Address;
                        var addressEnd = worksheet.Cell(currentRow - 1, col).Address;

                        if (col == 3) 
                            worksheet.Cell(currentRow, col).FormulaA1 = $"ROUND(SUM({addressStart}:{addressEnd}), 2)";
                        else
                            worksheet.Cell(currentRow, col).FormulaA1 = $"SUM({addressStart}:{addressEnd})";

                        worksheet.Cell(currentRow, col).Style.Font.Bold = true;

                        if (col == 3)
                            worksheet.Cell(currentRow, col).Style.NumberFormat.NumberFormatId = 2;
                        else
                            worksheet.Cell(currentRow, col).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }

                    worksheet.Columns().AdjustToContents();

                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_упаковщицы_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

                    workbook.SaveAs(tempFilePath);
                    Process.Start(new ProcessStartInfo(tempFilePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DateTime startDate = dpStartDate.SelectedDate.Value;
            DateTime endDate = dpEndDate.SelectedDate.Value;
            GeneratePackerReport(startDate, endDate);
        }
    }
}
