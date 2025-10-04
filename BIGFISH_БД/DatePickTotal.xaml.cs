using ClosedXML.Excel;
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
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для DatePickTotal.xaml
    /// </summary>
    public partial class DatePickTotal : Window
    {
        public DatePickTotal()
        {
            InitializeComponent();
        }

        private void GenerateCombinedReport_Click(object sender, RoutedEventArgs e)
        {
            DateTime startDate = dpStartDate.SelectedDate.Value;
            DateTime endDate = dpEndDate.SelectedDate.Value;

            string connectionString = @"data source=V_ISHENKA\SQLEXPRESS,1433;
            initial catalog=BigFishBD;
user id=User1;
password=12345;";

            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // ================== УПАКОВЩИЦЫ ==================
                    var packersSql = @"
WITH PackerStats AS (
    SELECT 
        p.FIO,
        SUM(ISNULL(dr.Packs, 0)) AS TotalPacks,
        SUM(ISNULL(dr.FinePacks, 0)) AS TotalFines,
        SUM(ISNULL(dr.Packs * a.PricePackers, 0)) AS Salary,
        SUM(ISNULL(dr.FinePacksFoundry * 2, 0)) AS Bonus
    FROM Packers p
    LEFT JOIN DailyReport dr ON p.FIO = dr.FIO 
        AND dr.DatePack BETWEEN @StartDate AND @EndDate
    LEFT JOIN Articles a ON dr.ArticlePack = a.Article
    GROUP BY p.FIO
),
DopUslugi AS (
    SELECT 
        dp.FIO,
        SUM(ISNULL(dp.Colvo * dp.PriceForOne, 0)) AS DopSum
    FROM DopPackers dp
    WHERE dp.DateDopPackers BETWEEN @StartDate AND @EndDate
    GROUP BY dp.FIO
),
AdvancePayments AS (
    SELECT 
        FIO,
        SUM(ISNULL(AdvancePay, 0)) AS TotalAdvance
    FROM AdvancePayPackers
    WHERE DateAdv BETWEEN @StartDate AND @EndDate
    GROUP BY FIO
)
SELECT 
    ps.FIO AS Упаковщица,
    ROUND(ps.TotalPacks, 2) AS [Кол-во пачек],
    ps.TotalFines AS [Штрафные пачки],
    ROUND(ps.TotalFines * 50, 2) AS [Сумма штрафа],
    ROUND(ps.Salary, 2) AS [Упаковка],
    ROUND(ps.Bonus, 2) AS [Доп плата за брак],
    ISNULL(du.DopSum, 0) AS [Допы],
    ISNULL(ap.TotalAdvance, 0) AS [АвансУдержания]
FROM PackerStats ps
LEFT JOIN DopUslugi du ON ps.FIO = du.FIO
LEFT JOIN AdvancePayments ap ON ps.FIO = ap.FIO
ORDER BY ps.FIO";

                    var packersTable = new DataTable();
                    using (var cmd = new SqlCommand(packersSql, connection))
                    {
                        cmd.Parameters.AddWithValue("@StartDate", startDate);
                        cmd.Parameters.AddWithValue("@EndDate", endDate);
                        packersTable.Load(cmd.ExecuteReader());
                    }

                    // Добавляем колонки для упаковщиц
                    packersTable.Columns.Add("Итого", typeof(decimal));

                    foreach (DataRow row in packersTable.Rows)
                    {
                        decimal salary = row["Упаковка"] != DBNull.Value ? Convert.ToDecimal(row["Упаковка"]) : 0m;
                        decimal finesAmount = row["Сумма штрафа"] != DBNull.Value ? Convert.ToDecimal(row["Сумма штрафа"]) : 0m;
                        decimal bonus = row["Доп плата за брак"] != DBNull.Value ? Convert.ToDecimal(row["Доп плата за брак"]) : 0m;
                        decimal dop = row["Допы"] != DBNull.Value ? Convert.ToDecimal(row["Допы"]) : 0m;
                        decimal advance = row["АвансУдержания"] != DBNull.Value ? Convert.ToDecimal(row["АвансУдержания"]) : 0m;
                        row["Итого"] = Math.Round(salary - finesAmount + bonus + dop + advance, 2);
                    }

                    // ================== ЛИТЕЙЩИКИ ==================
                    var foundersSql = @"
WITH TotalPacks AS (
    SELECT 
        f.FIO_Foundry,
        SUM(ISNULL(dr.Packs2, 0)) AS TotalPacks,
        SUM(ISNULL(dr.FinePacksFoundry, 0)) AS TotalFines
    FROM Foundry f
    LEFT JOIN DailyReport dr ON f.FIO_Foundry = dr.FIO_Foundry
        AND dr.DatePack BETWEEN @StartDate AND @EndDate
    GROUP BY f.FIO_Foundry
),
DopUslugi AS (
    SELECT 
        df.FIO_Foundry,
        SUM(ISNULL(df.Colvo * df.PriceForOne, 0)) AS DopSum
    FROM DopFoundry df
    WHERE df.DateDop BETWEEN @StartDate AND @EndDate
    GROUP BY df.FIO_Foundry
),
AdvancePayments AS (
    SELECT 
        FIO_Foundry,
        SUM(ISNULL(AdvancePay, 0)) AS TotalAdvance
    FROM AdvancePayFoundry
    WHERE DateAdv BETWEEN @StartDate AND @EndDate
    GROUP BY FIO_Foundry
)
SELECT 
    f.FIO_Foundry AS Литейщик,
    ROUND(SUM(ISNULL(dr.Packs2, 0)), 2) AS [Общее кол-во пачек],
    ROUND(SUM(ISNULL(CASE WHEN a.Type = 1 THEN dr.Packs2 ELSE 0 END, 0)), 2) AS [Стандартные пачки],
    ROUND(SUM(ISNULL(dr.Packs2 * a.PriceFoundry, 0)), 2) AS [Литьё],
    ROUND(SUM(ISNULL(dr.FinePacksFoundry, 0)), 2) AS [Пачки с браком],
    CASE
        WHEN ISNULL(tp.TotalFines, 0) > 0.05 * ISNULL(tp.TotalPacks, 1)
        THEN ROUND((ISNULL(tp.TotalFines, 0) - (0.05 * ISNULL(tp.TotalPacks, 1))) * 12, 2)
        ELSE 0
    END AS Штрафы,
    CASE 
        WHEN ISNULL(SUM(CASE WHEN a.Type = 1 THEN dr.Packs2 ELSE 0 END), 0) = 
             MAX(ISNULL(SUM(CASE WHEN a.Type = 1 THEN dr.Packs2 ELSE 0 END), 0)) OVER () 
        THEN ROUND(ISNULL(SUM(CASE WHEN a.Type = 1 THEN dr.Packs2 ELSE 0 END), 0) * 3, 2) 
        ELSE 0 
    END AS Премия,
    ISNULL(du.DopSum, 0) AS [Допы],
    ISNULL(ap.TotalAdvance, 0) AS [АвансУдержания]
FROM Foundry f
LEFT JOIN DailyReport dr ON f.FIO_Foundry = dr.FIO_Foundry 
    AND dr.DatePack BETWEEN @StartDate AND @EndDate
LEFT JOIN Articles a ON dr.ArticleFoundry = a.Article
LEFT JOIN TotalPacks tp ON f.FIO_Foundry = tp.FIO_Foundry
LEFT JOIN DopUslugi du ON f.FIO_Foundry = du.FIO_Foundry
LEFT JOIN AdvancePayments ap ON f.FIO_Foundry = ap.FIO_Foundry
GROUP BY f.FIO_Foundry, tp.TotalPacks, tp.TotalFines, du.DopSum, ap.TotalAdvance
ORDER BY f.FIO_Foundry";

                    var foundersTable = new DataTable();
                    using (var cmd = new SqlCommand(foundersSql, connection))
                    {
                        cmd.Parameters.AddWithValue("@StartDate", startDate);
                        cmd.Parameters.AddWithValue("@EndDate", endDate);
                        foundersTable.Load(cmd.ExecuteReader());
                    }

                    // Добавляем колонку Итого
                    foundersTable.Columns.Add("Итого", typeof(decimal));

                    foreach (DataRow row in foundersTable.Rows)
                    {
                        decimal premium = row["Премия"] != DBNull.Value
                            ? Convert.ToDecimal(row["Премия"])
                            : 0m;
                        decimal foundry = row["Литьё"] != DBNull.Value
                            ? Convert.ToDecimal(row["Литьё"])
                            : 0m;
                        decimal fines = row["Штрафы"] != DBNull.Value
                            ? Convert.ToDecimal(row["Штрафы"])
                            : 0m;
                        decimal dop = row["Допы"] != DBNull.Value
                            ? Convert.ToDecimal(row["Допы"])
                            : 0m;
                        decimal advance = row["АвансУдержания"] != DBNull.Value
                            ? Convert.ToDecimal(row["АвансУдержания"])
                            : 0m;

                        row["Итого"] = Math.Round(foundry + premium + dop - fines + advance, 2);
                    }

                    // ================== СОЗДАНИЕ EXCEL ==================
                    using (var workbook = new XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Отчёт");

                        // Стили
                        var headerStyle = workbook.Style;
                        headerStyle.Font.Bold = true;
                        headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerStyle.Fill.BackgroundColor = XLColor.Yellow;

                        // Заголовок
                        ws.Cell("A1").Value = "ВЕДОМОСТЬ ЗАРАБОТНОЙ ПЛАТЫ";
                        ws.Range("A1:J1").Merge().Style = headerStyle;

                        // Таблица упаковщиков
                        int rowNum = 3;
                        ws.Cell(rowNum, 1).Value = "ФИО Упаковщицы";
                        ws.Cell(rowNum, 2).Value = "Кол-во пачек";
                        ws.Cell(rowNum, 3).Value = "Штрафные пачки";
                        ws.Cell(rowNum, 4).Value = "Сумма штрафа";
                        ws.Cell(rowNum, 5).Value = "Доп. Оплата";
                        ws.Cell(rowNum, 6).Value = "Упаковка";
                        ws.Cell(rowNum, 7).Value = "Доп плата за брак";
                        ws.Cell(rowNum, 8).Value = "Аванс-удержания";
                        ws.Cell(rowNum, 9).Value = "Итого";
                        ws.Range(rowNum, 1, rowNum, 9).Style = headerStyle;

                        rowNum++;
                        foreach (DataRow pRow in packersTable.Rows)
                        {
                            ws.Cell(rowNum, 1).Value = pRow["Упаковщица"].ToString();
                            ws.Cell(rowNum, 2).Value = Convert.ToDecimal(pRow["Кол-во пачек"]);
                            ws.Cell(rowNum, 3).Value = Convert.ToDecimal(pRow["Штрафные пачки"]);
                            ws.Cell(rowNum, 4).Value = Convert.ToDecimal(pRow["Сумма штрафа"]);
                            ws.Cell(rowNum, 5).Value = Convert.ToDecimal(pRow["Допы"]);
                            ws.Cell(rowNum, 6).Value = Convert.ToDecimal(pRow["Упаковка"]);
                            ws.Cell(rowNum, 7).Value = Convert.ToDecimal(pRow["Доп плата за брак"]);
                            ws.Cell(rowNum, 8).Value = Convert.ToDecimal(pRow["АвансУдержания"]);
                            ws.Cell(rowNum, 9).Value = Convert.ToDecimal(pRow["Итого"]);
                            rowNum++;
                        }

                        // Итоговая строка упаковщиков
                        ws.Cell(rowNum, 1).Value = "Итого:";
                        ws.Range(rowNum, 1, rowNum, 9).Style = headerStyle;
                        ws.Cell(rowNum, 2).FormulaA1 = $"SUM(B4:B{rowNum - 1})";
                        ws.Cell(rowNum, 3).FormulaA1 = $"SUM(C4:C{rowNum - 1})";
                        ws.Cell(rowNum, 4).FormulaA1 = $"SUM(D4:D{rowNum - 1})";
                        ws.Cell(rowNum, 5).FormulaA1 = $"SUM(E4:E{rowNum - 1})";
                        ws.Cell(rowNum, 6).FormulaA1 = $"SUM(F4:F{rowNum - 1})";
                        ws.Cell(rowNum, 7).FormulaA1 = $"SUM(G4:G{rowNum - 1})";
                        ws.Cell(rowNum, 8).FormulaA1 = $"SUM(H4:H{rowNum - 1})";
                        ws.Cell(rowNum, 9).FormulaA1 = $"SUM(I4:I{rowNum - 1})";

                        // Таблица литейщиков
                        rowNum += 2;
                        ws.Cell(rowNum, 1).Value = "ФИО Литейщика";
                        ws.Cell(rowNum, 2).Value = "Стандартные пачки";
                        ws.Cell(rowNum, 3).Value = "Общее кол-во пачек";
                        ws.Cell(rowNum, 4).Value = "Доп. Оплата";
                        ws.Cell(rowNum, 5).Value = "Премия";
                        ws.Cell(rowNum, 6).Value = "Литьё";
                        ws.Cell(rowNum, 7).Value = "Штрафы";
                        ws.Cell(rowNum, 8).Value = "Аванс-удержания";
                        ws.Cell(rowNum, 9).Value = "Итого";
                        ws.Range(rowNum, 1, rowNum, 9).Style = headerStyle;

                        rowNum++;
                        foreach (DataRow fRow in foundersTable.Rows)
                        {
                            ws.Cell(rowNum, 1).Value = fRow["Литейщик"].ToString();
                            ws.Cell(rowNum, 2).Value = Convert.ToDecimal(fRow["Стандартные пачки"]);
                            ws.Cell(rowNum, 3).Value = Convert.ToDecimal(fRow["Общее кол-во пачек"]);
                            ws.Cell(rowNum, 4).Value = Convert.ToDecimal(fRow["Допы"]);
                            ws.Cell(rowNum, 5).Value = Convert.ToDecimal(fRow["Премия"]);
                            ws.Cell(rowNum, 6).Value = Convert.ToDecimal(fRow["Литьё"]);
                            ws.Cell(rowNum, 7).Value = Convert.ToDecimal(fRow["Штрафы"]);
                            ws.Cell(rowNum, 8).Value = Convert.ToDecimal(fRow["АвансУдержания"]);
                            ws.Cell(rowNum, 9).Value = Convert.ToDecimal(fRow["Итого"]);
                            rowNum++;
                        }

                        // Итоговая строка литейщиков
                        ws.Cell(rowNum, 1).Value = "Итого:";
                        ws.Range(rowNum, 1, rowNum, 9).Style = headerStyle;
                        ws.Cell(rowNum, 2).FormulaA1 = $"SUM(B{rowNum - foundersTable.Rows.Count}:B{rowNum - 1})";
                        ws.Cell(rowNum, 3).FormulaA1 = $"SUM(C{rowNum - foundersTable.Rows.Count}:C{rowNum - 1})";
                        ws.Cell(rowNum, 4).FormulaA1 = $"SUM(D{rowNum - foundersTable.Rows.Count}:D{rowNum - 1})";
                        ws.Cell(rowNum, 5).FormulaA1 = $"SUM(E{rowNum - foundersTable.Rows.Count}:E{rowNum - 1})";
                        ws.Cell(rowNum, 6).FormulaA1 = $"SUM(F{rowNum - foundersTable.Rows.Count}:F{rowNum - 1})";
                        ws.Cell(rowNum, 7).FormulaA1 = $"SUM(G{rowNum - foundersTable.Rows.Count}:G{rowNum - 1})";
                        ws.Cell(rowNum, 8).FormulaA1 = $"SUM(H{rowNum - foundersTable.Rows.Count}:H{rowNum - 1})";
                        ws.Cell(rowNum, 9).FormulaA1 = $"SUM(I{rowNum - foundersTable.Rows.Count}:I{rowNum - 1})";

                        // Форматирование
                        ws.Columns().AdjustToContents();
                        ws.Range("B4:C" + rowNum).Style.NumberFormat.NumberFormatId = 2;
                        ws.Range("D4:J" + rowNum).Style.NumberFormat.Format = "#,##0.00 ₽";

                        // Сохранение
                        var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                            $"Итоговый_отчет{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

                        workbook.SaveAs(tempFilePath);
                        Process.Start(new ProcessStartInfo(tempFilePath) { UseShellExecute = true });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}\n\n{ex.StackTrace}");
            }
        }
    }
}
