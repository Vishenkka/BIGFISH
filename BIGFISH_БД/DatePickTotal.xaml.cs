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

            try
            {
                using (var db = new BigFishBDEntities())
                {
                    //УПАКОВЩИЦЫ
                    var packersData = db.Packers
                        .Select(p => new
                        {
                            FIO = p.FIO,
                            DailyReports = db.DailyReport
                                .Where(dr => dr.FIO == p.FIO && dr.DatePack >= startDate && dr.DatePack <= endDate)
                                .ToList(),
                            DopUslugi = db.DopPackers
                                .Where(dp => dp.FIO == p.FIO && dp.DateDopPackers >= startDate && dp.DateDopPackers <= endDate)
                                .ToList(),
                            AdvancePayments = db.AdvancePayPackers
                                .Where(ap => ap.FIO == p.FIO && ap.DateAdv >= startDate && ap.DateAdv <= endDate)
                                .ToList()
                        })
                        .ToList()
                        .Select(p => new
                        {
                            Упаковщица = p.FIO,
                            TotalPacks = p.DailyReports.Sum(dr => dr.Packs ?? 0),
                            TotalFines = p.DailyReports.Sum(dr => dr.FinePacks ?? 0),
                            Salary = p.DailyReports.Sum(dr => (dr.Packs ?? 0) * (dr.Articles?.PricePackers ?? 0)),
                            Bonus = p.DailyReports.Sum(dr => (dr.FinePacksFoundry ?? 0) * 2),
                            DopSum = p.DopUslugi.Sum(dp => (dp.Colvo ?? 0) * (dp.PriceForOne ?? 0)),
                            TotalAdvance = p.AdvancePayments.Sum(ap => ap.AdvancePay ?? 0)
                        })
                        .Select(p => new
                        {
                            Упаковщица = p.Упаковщица,
                            КоличествоПачек = Math.Round(p.TotalPacks, 2),
                            ШтрафныеПачки = p.TotalFines,
                            СуммаШтрафа = Math.Round(p.TotalFines * 50, 2),
                            Упаковка = Math.Round(p.Salary, 2),
                            ДопПлатаЗаБрак = Math.Round(p.Bonus, 2),
                            Допы = p.DopSum,
                            АвансУдержания = p.TotalAdvance,
                            Итого = Math.Round(p.Salary - (p.TotalFines * 50) + p.Bonus + p.DopSum + p.TotalAdvance, 2)
                        })
                        .OrderBy(p => p.Упаковщица)
                        .ToList();

                    //ЛИТЕЙЩИКИ
                    var foundersData = db.Foundry
                        .Select(f => new
                        {
                            FIO = f.FIO_Foundry,
                            DailyReports = db.DailyReport
                                .Where(dr => dr.FIO_Foundry == f.FIO_Foundry && dr.DatePack >= startDate && dr.DatePack <= endDate)
                                .ToList(),
                            DopUslugi = db.DopFoundry
                                .Where(df => df.FIO_Foundry == f.FIO_Foundry && df.DateDop >= startDate && df.DateDop <= endDate)
                                .ToList(),
                            AdvancePayments = db.AdvancePayFoundry
                                .Where(af => af.FIO_Foundry == f.FIO_Foundry && af.DateAdv >= startDate && af.DateAdv <= endDate)
                                .ToList()
                        })
                        .ToList()
                        .Select(f => new
                        {
                            FIO = f.FIO,
                            TotalPacks = f.DailyReports.Sum(dr => dr.Packs2 ?? 0),
                            TotalFines = f.DailyReports.Sum(dr => dr.FinePacksFoundry ?? 0),
                            StandardPacks = f.DailyReports.Where(dr => dr.Articles?.Type == 1).Sum(dr => dr.Packs2 ?? 0),
                            FoundrySalary = f.DailyReports.Sum(dr => (dr.Packs2 ?? 0) * (dr.Articles?.PriceFoundry ?? 0)),
                            DopSum = f.DopUslugi.Sum(df => (df.Colvo) * (df.PriceForOne)),
                            TotalAdvance = f.AdvancePayments.Sum(af => af.AdvancePay ?? 0)
                        })
                        .ToList();



                    var maxStandardPacks = foundersData.Max(f => (decimal)f.StandardPacks);

                    var foundersReport = foundersData
                        .Select(f => new
                        {
                            Литейщик = f.FIO,
                            ОбщееКоличествоПачек = Math.Round((decimal)f.TotalPacks, 2),
                            СтандартныеПачки = Math.Round((decimal)f.StandardPacks, 2),
                            Литьё = Math.Round((decimal)f.FoundrySalary, 2),
                            ПачкиСБраком = Math.Round((decimal)f.TotalFines, 2),
                            Штрафы = (decimal)f.TotalFines > (0.05m * ((decimal)f.TotalPacks == 0 ? 1 : (decimal)f.TotalPacks))
                                ? Math.Round(((decimal)f.TotalFines - (0.05m * (decimal)f.TotalPacks)) * 12, 2)
                                : 0m,
                            Премия = (decimal)f.StandardPacks == maxStandardPacks && maxStandardPacks > 0
                                ? Math.Round((decimal)f.StandardPacks * 3, 2)
                                : 0m,
                            Допы = (decimal)f.DopSum,
                            АвансУдержания = (decimal)f.TotalAdvance,
                            Итого = Math.Round((decimal)f.FoundrySalary +
                                ((decimal)f.StandardPacks == maxStandardPacks && maxStandardPacks > 0 ? (decimal)f.StandardPacks * 3 : 0) +
                                (decimal)f.DopSum -
                                ((decimal)f.TotalFines > (0.05m * ((decimal)f.TotalPacks == 0 ? 1 : (decimal)f.TotalPacks)) ?
                                    ((decimal)f.TotalFines - (0.05m * (decimal)f.TotalPacks)) * 12 : 0) +
                                (decimal)f.TotalAdvance, 2)
                        })
                        .OrderBy(f => f.Литейщик)
                        .ToList();



                    using (var workbook = new XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Отчёт");

                     
                        var headerStyle = workbook.Style;
                        headerStyle.Font.Bold = true;
                        headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerStyle.Fill.BackgroundColor = XLColor.Yellow;

                        ws.Cell("A1").Value = "ВЕДОМОСТЬ ЗАРАБОТНОЙ ПЛАТЫ";
                        ws.Range("A1:J1").Merge().Style = headerStyle;


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
                        foreach (var packer in packersData)
                        {
                            ws.Cell(rowNum, 1).Value = packer.Упаковщица;
                            ws.Cell(rowNum, 2).Value = packer.КоличествоПачек;
                            ws.Cell(rowNum, 3).Value = packer.ШтрафныеПачки;
                            ws.Cell(rowNum, 4).Value = packer.СуммаШтрафа;
                            ws.Cell(rowNum, 5).Value = packer.Допы;
                            ws.Cell(rowNum, 6).Value = packer.Упаковка;
                            ws.Cell(rowNum, 7).Value = packer.ДопПлатаЗаБрак;
                            ws.Cell(rowNum, 8).Value = packer.АвансУдержания;
                            ws.Cell(rowNum, 9).Value = packer.Итого;
                            rowNum++;
                        }


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
                        foreach (var founder in foundersReport)
                        {
                            ws.Cell(rowNum, 1).Value = founder.Литейщик;
                            ws.Cell(rowNum, 2).Value = founder.СтандартныеПачки;
                            ws.Cell(rowNum, 3).Value = founder.ОбщееКоличествоПачек;
                            ws.Cell(rowNum, 4).Value = founder.Допы;
                            ws.Cell(rowNum, 5).Value = founder.Премия;
                            ws.Cell(rowNum, 6).Value = founder.Литьё;
                            ws.Cell(rowNum, 7).Value = founder.Штрафы;
                            ws.Cell(rowNum, 8).Value = founder.АвансУдержания;
                            ws.Cell(rowNum, 9).Value = founder.Итого;
                            rowNum++;
                        }


                        ws.Cell(rowNum, 1).Value = "Итого:";
                        ws.Range(rowNum, 1, rowNum, 9).Style = headerStyle;

                        int foundersStartRow = rowNum - foundersReport.Count;
                        int foundersEndRow = rowNum - 1;

                        ws.Cell(rowNum, 2).FormulaA1 = $"SUM(B{foundersStartRow}:B{foundersEndRow})";
                        ws.Cell(rowNum, 3).FormulaA1 = $"SUM(C{foundersStartRow}:C{foundersEndRow})";
                        ws.Cell(rowNum, 4).FormulaA1 = $"SUM(D{foundersStartRow}:D{foundersEndRow})";
                        ws.Cell(rowNum, 5).FormulaA1 = $"SUM(E{foundersStartRow}:E{foundersEndRow})";
                        ws.Cell(rowNum, 6).FormulaA1 = $"SUM(F{foundersStartRow}:F{foundersEndRow})";
                        ws.Cell(rowNum, 7).FormulaA1 = $"SUM(G{foundersStartRow}:G{foundersEndRow})";
                        ws.Cell(rowNum, 8).FormulaA1 = $"SUM(H{foundersStartRow}:H{foundersEndRow})";
                        ws.Cell(rowNum, 9).FormulaA1 = $"SUM(I{foundersStartRow}:I{foundersEndRow})";


                        ws.Columns().AdjustToContents();
                        ws.Range("B4:C" + rowNum).Style.NumberFormat.NumberFormatId = 2;
                        ws.Range("D4:J" + rowNum).Style.NumberFormat.Format = "#,##0.00 ₽";

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
