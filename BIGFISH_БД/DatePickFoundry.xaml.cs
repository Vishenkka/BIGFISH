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

                    var orangeStyle = workbook.Style;
                    orangeStyle.Fill.BackgroundColor = XLColor.Yellow;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽;-#,##0.00 ₽";

                    //ТАБЛИЦА 1 Детализация по артикулам 
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

                    //ТАБЛИЦА 2 Сводные данные
                    var sql = @"
                                WITH AllFounders AS (
                                    SELECT DISTINCT FIO_Foundry FROM Foundry
                                    WHERE FIO_Foundry IN (
                                        SELECT FIO_Foundry FROM DailyReport WHERE DatePack BETWEEN @start AND @end
                                        UNION
                                        SELECT FIO_Foundry FROM DopFoundry WHERE DateDop BETWEEN @start AND @end
                                        UNION
                                        SELECT FIO_Foundry FROM AdvancePayFoundry WHERE DateAdv BETWEEN @start AND @end
                                    )
                                ),
                                DailyReportData AS (
                                    SELECT 
                                        f.FIO_Foundry,
                                        SUM(ISNULL(dr.Packs2, 0)) AS TotalPacks,
                                        SUM(ISNULL(dr.FinePacksFoundry, 0)) AS TotalFines,
                                        SUM(ISNULL(dr.Packs2 * a.PriceFoundry, 0)) AS Литье,
                                        SUM(ISNULL(CASE WHEN a.Type = 1 THEN dr.Packs2 ELSE 0 END, 0)) AS СтандартныеПачки,
                                        SUM(ISNULL(dr.Packs2, 0)) AS ОбщееКолво,
                                        SUM(ISNULL(dr.FinePacksFoundry, 0)) AS ПачкиСБраком
                                    FROM AllFounders f
                                    LEFT JOIN DailyReport dr ON f.FIO_Foundry = dr.FIO_Foundry AND dr.DatePack BETWEEN @start AND @end
                                    LEFT JOIN Articles a ON dr.ArticleFoundry = a.Article
                                    GROUP BY f.FIO_Foundry
                                ),
                                DopUslugi AS (
                                    SELECT 
                                        f.FIO_Foundry,
                                        SUM(ISNULL(df.Colvo * df.PriceForOne, 0)) AS DopSum
                                    FROM AllFounders f
                                    LEFT JOIN DopFoundry df ON f.FIO_Foundry = df.FIO_Foundry AND df.DateDop BETWEEN @start AND @end
                                    GROUP BY f.FIO_Foundry
                                ),
                                AdvancePayments AS (
                                    SELECT 
                                        f.FIO_Foundry,
                                        SUM(ISNULL(ap.AdvancePay, 0)) AS TotalAdvance
                                    FROM AllFounders f
                                    LEFT JOIN AdvancePayFoundry ap ON f.FIO_Foundry = ap.FIO_Foundry AND ap.DateAdv BETWEEN @start AND @end
                                    GROUP BY f.FIO_Foundry
                                ),
                                TotalPacks AS (
                                    SELECT 
                                        FIO_Foundry,
                                        TotalPacks,
                                        TotalFines
                                    FROM DailyReportData
                                )
                                SELECT 
                                    f.FIO_Foundry AS Литейщик,
                                    ROUND(ISNULL(drd.Литье, 0), 2) AS [Литье],
                                    ROUND(ISNULL(drd.СтандартныеПачки, 0), 2) AS [Стандартные пачки],
                                    ROUND(ISNULL(drd.ОбщееКолво, 0), 2) AS [Общее кол-во],
                                    ROUND(ISNULL(drd.ПачкиСБраком, 0), 2) AS [Пачки с браком],
                                    CASE
                                        WHEN ISNULL(tp.TotalFines, 0) > 0.05 * ISNULL(tp.TotalPacks, 1)
                                        THEN ROUND((ISNULL(tp.TotalFines, 0) - (0.05 * ISNULL(tp.TotalPacks, 1))) * 12, 2)
                                        ELSE 0
                                    END AS Штрафы,
                                    CASE 
                                        WHEN ISNULL(drd.СтандартныеПачки, 0) = 
                                             MAX(ISNULL(drd.СтандартныеПачки, 0)) OVER () 
                                        THEN ROUND(ISNULL(drd.СтандартныеПачки, 0) * 3, 2) 
                                        ELSE 0 
                                    END AS Премия,
                                    ISNULL(du.DopSum, 0) AS [Допы],
                                    ISNULL(ap.TotalAdvance, 0) AS [АвансУдержания]
                                FROM AllFounders f
                                LEFT JOIN DailyReportData drd ON f.FIO_Foundry = drd.FIO_Foundry
                                LEFT JOIN TotalPacks tp ON f.FIO_Foundry = tp.FIO_Foundry
                                LEFT JOIN DopUslugi du ON f.FIO_Foundry = du.FIO_Foundry
                                LEFT JOIN AdvancePayments ap ON f.FIO_Foundry = ap.FIO_Foundry
                                ORDER BY f.FIO_Foundry";

                    var summaryHeaders = new[] { "Литейщик", "Литье", "Стандартные пачки", "Общее кол-во", "Пачки с браком", "Штрафы", "Премия", "Допы", "Аванс-удержания", "Итого" };
                    for (int i = 0; i < summaryHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = summaryHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;

                    var reports = new List<dynamic>();
                    using (var connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        using (var cmd = new SqlCommand(sql, connection))
                        {
                            cmd.Parameters.AddWithValue("@start", startDate);
                            cmd.Parameters.AddWithValue("@end", endDate);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    reports.Add(new
                                    {
                                        Литейщик = reader["Литейщик"].ToString(),
                                        Литье = reader["Литье"] != DBNull.Value ? Convert.ToDecimal(reader["Литье"]) : 0m,
                                        Стандартные = reader["Стандартные пачки"] != DBNull.Value ? Convert.ToDecimal(reader["Стандартные пачки"]) : 0m,
                                        ОбщееКолво = reader["Общее кол-во"] != DBNull.Value ? Convert.ToDecimal(reader["Общее кол-во"]) : 0m,
                                        Брак = reader["Пачки с браком"] != DBNull.Value ? Convert.ToDecimal(reader["Пачки с браком"]) : 0m,
                                        Штрафы = reader["Штрафы"] != DBNull.Value ? Convert.ToDecimal(reader["Штрафы"]) : 0m,
                                        Премия = reader["Премия"] != DBNull.Value ? Convert.ToDecimal(reader["Премия"]) : 0m,
                                        Допы = reader["Допы"] != DBNull.Value ? Convert.ToDecimal(reader["Допы"]) : 0m,
                                        АвансУдержания = reader["АвансУдержания"] != DBNull.Value ? Convert.ToDecimal(reader["АвансУдержания"]) : 0m
                                    });
                                }
                            }
                        }
                    }

                    foreach (var report in reports)
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
                        var addressStart = worksheet.Cell(currentRow - reports.Count, col).Address;
                        var addressEnd = worksheet.Cell(currentRow - 1, col).Address;

                        if (col == 3 || col == 4 || col == 5) // Для числовых колонок
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







       







