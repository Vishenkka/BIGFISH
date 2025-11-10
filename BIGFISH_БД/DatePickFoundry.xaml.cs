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
using DocumentFormat.OpenXml.ExtendedProperties;
using System.IO;
using System.Data.Entity;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.InteropServices.ComTypes;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;


namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для DatePickFoundry.xaml
    /// </summary>
    public partial class DatePickFoundry : Window
    {
        private readonly BigFishBDEntities _dbContext;
        public DatePickFoundry()
        {
            InitializeComponent();
            _dbContext = new BigFishBDEntities();
        }


        private void GenerateCombinedFoundryReport(DateTime startDate, DateTime endDate)
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
                    var articleDetails = db.Foundry
                        .Select(f => new
                        {
                            Литейщик = f.FIO_Foundry,
                            Данные = db.DailyReport
                                .Where(dr => dr.FIO_Foundry == f.FIO_Foundry &&
                                            dr.DatePack >= startDate &&
                                            dr.DatePack <= endDate)
                                .GroupBy(dr => dr.ArticleFoundry)
                                .Select(g => new
                                {
                                    Article = g.Key,
                                    Пачки = g.Sum(x => x.Packs2 ?? 0)
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
                        .ToDictionary(a => a.Article, a => a.PriceFoundry);

                    int currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Литейщик";
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
                        worksheet.Cell(currentRow, 1).Value = detail.Литейщик;

                        var articleDict = detail.Данные
                            .GroupBy(d => d.Article)
                            .ToDictionary(g => g.Key, g => Math.Round(g.Sum(x => x.Пачки), 2));

                        for (int i = 0; i < allArticles.Count; i++)
                        {
                            var article = allArticles[i];
                            var value = articleDict.ContainsKey(article) ? articleDict[article] : 0;
                            worksheet.Cell(currentRow, i + 2).Value = value;
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
                    var allFounders = db.Foundry
                        .Where(f => db.DailyReport.Any(dr => dr.FIO_Foundry == f.FIO_Foundry && dr.DatePack >= startDate && dr.DatePack <= endDate) ||
                                   db.DopFoundry.Any(df => df.FIO_Foundry == f.FIO_Foundry && df.DateDop >= startDate && df.DateDop <= endDate) ||
                                   db.AdvancePayFoundry.Any(ap => ap.FIO_Foundry == f.FIO_Foundry && ap.DateAdv >= startDate && ap.DateAdv <= endDate))
                        .Select(f => f.FIO_Foundry)
                        .Distinct()
                        .ToList();

                    var allDailyData = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .ToList();

                    var allDopData = db.DopFoundry
                        .Where(df => df.DateDop >= startDate && df.DateDop <= endDate)
                        .ToList();

                    var allAdvanceData = db.AdvancePayFoundry
                        .Where(ap => ap.DateAdv >= startDate && ap.DateAdv <= endDate)
                        .ToList();


                    var dailyReportData = allDailyData
                        .GroupBy(d => d.FIO_Foundry)
                        .Select(g => new
                        {
                            FIO_Foundry = g.Key,
                            TotalPacks = g.Sum(x => (decimal)(x.Packs2 ?? 0)),
                            TotalFines = g.Sum(x => (decimal)(x.FinePacksFoundry ?? 0)),
                            Литье = g.Sum(x => (decimal)(x.Packs2 ?? 0) * (x.Articles?.PriceFoundry ?? 0)),
                            СтандартныеПачки = g.Where(x => x.Articles?.Type == 1).Sum(x => (decimal)(x.Packs2 ?? 0)),
                            ОбщееКолво = g.Sum(x => (decimal)(x.Packs2 ?? 0)),
                            ПачкиСБраком = g.Sum(x => (decimal)(x.FinePacksFoundry ?? 0))
                        })
                        .ToList();


                    var dopUslugiData = allDopData
                        .GroupBy(d => d.FIO_Foundry)
                        .Select(g => new
                        {
                            FIO_Foundry = g.Key,
                            DopSum = g.Sum(x => (decimal)(x.Colvo) * (decimal)(x.PriceForOne))
                        })
                        .ToList();


                    var advancePaymentsData = allAdvanceData
                        .GroupBy(a => a.FIO_Foundry)
                        .Select(g => new
                        {
                            FIO_Foundry = g.Key,
                            TotalAdvance = g.Sum(x => (decimal)(x.AdvancePay ?? 0))
                        })
                        .ToList();


                    decimal maxStandardPacks = 0m;
                    if (dailyReportData.Any())
                    {
                        maxStandardPacks = dailyReportData.Max(x => x.СтандартныеПачки);
                    }


                    var reportsList = allFounders
                        .Select(f =>
                        {
                            var dailyData = dailyReportData.FirstOrDefault(d => d.FIO_Foundry == f);
                            var dopData = dopUslugiData.FirstOrDefault(d => d.FIO_Foundry == f);
                            var advanceData = advancePaymentsData.FirstOrDefault(a => a.FIO_Foundry == f);

                            decimal totalPacks = dailyData != null ? dailyData.TotalPacks : 0m;
                            decimal totalFines = dailyData != null ? dailyData.TotalFines : 0m;
                            decimal литье = dailyData != null ? dailyData.Литье : 0m;
                            decimal стандартныеПачки = dailyData != null ? dailyData.СтандартныеПачки : 0m;
                            decimal общееКолво = dailyData != null ? dailyData.ОбщееКолво : 0m;
                            decimal пачкиСБраком = dailyData != null ? dailyData.ПачкиСБраком : 0m;
                            decimal допы = dopData != null ? dopData.DopSum : 0m;
                            decimal авансУдержания = advanceData != null ? advanceData.TotalAdvance : 0m;


                            decimal штрафы = 0m;
                            if (totalFines > 0.05m * (totalPacks > 0 ? totalPacks : 1))
                            {
                                штрафы = Math.Round((totalFines - (0.05m * totalPacks)) * 12, 2);
                            }


                            decimal премия = 0m;
                            if (стандартныеПачки > 0 && Math.Abs(стандартныеПачки - maxStandardPacks) < 0.001m)
                            {
                                премия = Math.Round(стандартныеПачки * 3, 2);
                            }

                            return new
                            {
                                Литейщик = f,
                                Литье = Math.Round(литье, 2),
                                Стандартные = Math.Round(стандартныеПачки, 2),
                                ОбщееКолво = Math.Round(общееКолво, 2),
                                Брак = Math.Round(пачкиСБраком, 2),
                                Штрафы = штрафы,
                                Премия = премия,
                                Допы = допы,
                                АвансУдержания = авансУдержания
                            };
                        })
                        .OrderBy(r => r.Литейщик)
                        .ToList();

                    var summaryHeaders = new[] { "Литейщик", "Литье", "Стандартные пачки", "Общее кол-во", "Пачки с браком", "Штрафы", "Премия", "Допы", "Аванс-удержания", "Итого" };
                    for (int i = 0; i < summaryHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = summaryHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;

                    foreach (var report in reportsList)
                    {
                        worksheet.Cell(currentRow, 1).Value = report.Литейщик;
                        worksheet.Cell(currentRow, 2).Value = report.Литье;
                        worksheet.Cell(currentRow, 2).Style = moneyStyle;
                        worksheet.Cell(currentRow, 3).Value = report.Стандартные;
                        worksheet.Cell(currentRow, 3).Style.NumberFormat.NumberFormatId = 2;
                        worksheet.Cell(currentRow, 4).Value = report.ОбщееКолво;
                        worksheet.Cell(currentRow, 4).Style.NumberFormat.NumberFormatId = 2;
                        worksheet.Cell(currentRow, 5).Value = report.Брак;
                        worksheet.Cell(currentRow, 5).Style.NumberFormat.NumberFormatId = 2;
                        worksheet.Cell(currentRow, 6).Value = report.Штрафы;
                        worksheet.Cell(currentRow, 6).Style = moneyStyle;
                        worksheet.Cell(currentRow, 7).Value = report.Премия;
                        worksheet.Cell(currentRow, 7).Style = moneyStyle;
                        worksheet.Cell(currentRow, 8).Value = report.Допы;
                        worksheet.Cell(currentRow, 8).Style = moneyStyle;
                        worksheet.Cell(currentRow, 9).Value = report.АвансУдержания;
                        worksheet.Cell(currentRow, 9).Style = moneyStyle;
                        worksheet.Cell(currentRow, 10).Value = report.Литье + report.Премия + report.Допы - report.Штрафы + report.АвансУдержания;
                        worksheet.Cell(currentRow, 10).Style = moneyStyle;
                        currentRow++;
                    }

                    worksheet.Cell(currentRow, 1).Value = "Итого:";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;

                    for (int col = 2; col <= 10; col++)
                    {
                        int startRowNum = currentRow - reportsList.Count;
                        int endRowNum = currentRow - 1;

                        var addressStart = worksheet.Cell(startRowNum, col).Address;
                        var addressEnd = worksheet.Cell(endRowNum, col).Address;

                        if (col == 3 || col == 4 || col == 5) 
                            worksheet.Cell(currentRow, col).FormulaA1 = $"ROUND(SUM({addressStart}:{addressEnd}), 2)";
                        else
                            worksheet.Cell(currentRow, col).FormulaA1 = $"SUM({addressStart}:{addressEnd})";

                        worksheet.Cell(currentRow, col).Style.Font.Bold = true;

                        if (col == 3 || col == 4 || col == 5)
                            worksheet.Cell(currentRow, col).Style.NumberFormat.NumberFormatId = 2;
                        else
                            worksheet.Cell(currentRow, col).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }

                    worksheet.Columns().AdjustToContents();

                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_литейщики_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

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
            GenerateCombinedFoundryReport(startDate, endDate);
        }
    }
}







       







