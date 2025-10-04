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

                    // Стили оформления
                    var headerStyle = workbook.Style;
                    headerStyle.Font.Bold = true;
                    headerStyle.Fill.BackgroundColor = XLColor.Yellow;
                    headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var orangeStyle = workbook.Style;
                    orangeStyle.Fill.BackgroundColor = XLColor.Yellow;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽;-#,##0.00 ₽";

                    // ===== ТАБЛИЦА 1: Детализация по артикулам =====
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

                    // Заголовки таблицы 1
                    int currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Упаковщица";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;

                    for (int i = 0; i < allArticles.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 2).Value = allArticles[i];
                        worksheet.Cell(currentRow, i + 2).Style = headerStyle;
                    }

                    // Строка с ценами
                    currentRow++;
                    worksheet.Row(currentRow).Style = orangeStyle;
                    for (int i = 0; i < allArticles.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 2).Value = articlePrices[allArticles[i]];
                        worksheet.Cell(currentRow, i + 2).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }

                    // Данные таблицы 1 (с округлением)
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

                    // Итоговая строка по пачкам (числовой формат с округлением)
                    worksheet.Cell(currentRow, 1).Value = "Итого:";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;
                    for (int col = 2; col <= allArticles.Count + 1; col++)
                    {
                        // Округляем результат формулы до сотых
                        worksheet.Cell(currentRow, col).FormulaA1 = $"ROUND(SUM({worksheet.Cell(3, col).Address}:{worksheet.Cell(currentRow - 1, col).Address}), 2)";
                        worksheet.Cell(currentRow, col).Style = headerStyle;
                        worksheet.Cell(currentRow, col).Style.NumberFormat.NumberFormatId = 2;
                    }

                    // Итоговая строка по сумме (денежный формат)
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

                    // ===== ТАБЛИЦА 2: Сводные данные =====
                    var sql = @"
WITH AllPackers AS (
    SELECT DISTINCT FIO FROM Packers
    WHERE FIO IN (
        SELECT FIO FROM DailyReport WHERE DatePack BETWEEN @start AND @end
        UNION
        SELECT FIO FROM AdvancePayPackers WHERE DateAdv BETWEEN @start AND @end
        UNION
        SELECT FIO FROM DopPackers WHERE DateDopPackers BETWEEN @start AND @end
    )
),
DailyReportData AS (
    SELECT 
        p.FIO,
        SUM(ISNULL(dr.Packs, 0)) AS ВсегоПачек,
        SUM(ISNULL(dr.FinePacks, 0)) AS ВсегоШтрафов,
        SUM(ISNULL(dr.FinePacksFoundry, 0)) * 2 AS Премия,
        SUM(ISNULL(dr.Packs * a.PricePackers, 0)) AS Упаковка
    FROM AllPackers p
    LEFT JOIN DailyReport dr ON p.FIO = dr.FIO AND dr.DatePack BETWEEN @start AND @end
    LEFT JOIN Articles a ON dr.ArticlePack = a.Article
    GROUP BY p.FIO
),
DopUslugi AS (
    SELECT 
        p.FIO,
        SUM(ISNULL(dp.Colvo * dp.PriceForOne, 0)) AS DopSum
    FROM AllPackers p
    LEFT JOIN DopPackers dp ON p.FIO = dp.FIO AND dp.DateDopPackers BETWEEN @start AND @end
    GROUP BY p.FIO
),
AdvancePayments AS (
    SELECT 
        p.FIO,
        SUM(ISNULL(ap.AdvancePay, 0)) AS TotalAdvance
    FROM AllPackers p
    LEFT JOIN AdvancePayPackers ap ON p.FIO = ap.FIO AND ap.DateAdv BETWEEN @start AND @end
    GROUP BY p.FIO
)
SELECT 
    p.FIO AS Упаковщица,
    ROUND(ISNULL(drd.ВсегоПачек, 0), 2) AS Пачки,
    ROUND(ISNULL(drd.ВсегоШтрафов, 0) * 50, 2) AS Штрафы,
    ISNULL(drd.Премия, 0) AS Премия,
    ISNULL(drd.Упаковка, 0) AS Упаковка,
    ISNULL(du.DopSum, 0) AS Допы,
    ISNULL(ap.TotalAdvance, 0) AS АвансУдержания
FROM AllPackers p
LEFT JOIN DailyReportData drd ON p.FIO = drd.FIO
LEFT JOIN DopUslugi du ON p.FIO = du.FIO
LEFT JOIN AdvancePayments ap ON p.FIO = ap.FIO
ORDER BY p.FIO";

                    // Заголовки таблицы 2
                    var summaryHeaders = new[] { "Упаковщица", "Упаковка", "Количество пачек", "Штрафы", "Премия", "Допы", "Аванс-удержания", "Итого" };
                    for (int i = 0; i < summaryHeaders.Length; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = summaryHeaders[i];
                        worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                    }
                    currentRow++;

                    // Данные таблицы 2
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
                                        Упаковщица = reader["Упаковщица"].ToString(),
                                        Пачки = reader["Пачки"] != DBNull.Value ? Convert.ToDecimal(reader["Пачки"]) : 0m,
                                        Штрафы = reader["Штрафы"] != DBNull.Value ? Convert.ToDecimal(reader["Штрафы"]) : 0m,
                                        Премия = reader["Премия"] != DBNull.Value ? Convert.ToDecimal(reader["Премия"]) : 0m,
                                        Упаковка = reader["Упаковка"] != DBNull.Value ? Convert.ToDecimal(reader["Упаковка"]) : 0m,
                                        Допы = reader["Допы"] != DBNull.Value ? Convert.ToDecimal(reader["Допы"]) : 0m,
                                        АвансУдержания = reader["АвансУдержания"] != DBNull.Value ? Convert.ToDecimal(reader["АвансУдержания"]) : 0m
                                    });
                                }
                            }
                        }
                    }

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

                    // Итоговая строка таблицы 2
                    worksheet.Cell(currentRow, 1).Value = "Итого:";
                    worksheet.Cell(currentRow, 1).Style = headerStyle;

                    for (int col = 2; col <= 8; col++)
                    {
                        var addressStart = worksheet.Cell(currentRow - reports.Count, col).Address;
                        var addressEnd = worksheet.Cell(currentRow - 1, col).Address;

                        if (col == 3) // Для "Количество пачек" - округляем
                            worksheet.Cell(currentRow, col).FormulaA1 = $"ROUND(SUM({addressStart}:{addressEnd}), 2)";
                        else
                            worksheet.Cell(currentRow, col).FormulaA1 = $"SUM({addressStart}:{addressEnd})";

                        worksheet.Cell(currentRow, col).Style.Font.Bold = true;

                        if (col == 3)
                            worksheet.Cell(currentRow, col).Style.NumberFormat.NumberFormatId = 2;
                        else
                            worksheet.Cell(currentRow, col).Style.NumberFormat.Format = moneyStyle.NumberFormat.Format;
                    }

                    // Финализация документа
                    worksheet.Columns().AdjustToContents();

                    // Сохранение во временный файл и открытие
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
