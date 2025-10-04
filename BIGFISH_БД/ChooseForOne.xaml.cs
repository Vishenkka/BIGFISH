using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
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
    /// Логика взаимодействия для ChooseForOne.xaml
    /// </summary>
    public partial class ChooseForOne : UserControl
    {
        BigFishBDEntities db;
        public List<Packers> packers { get; set; }

        public List<Foundry> foundry { get; set; }

        public ChooseForOne()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            LoadPackersData();
            LoadFoundryData();

        }

        #region кнопки и загрузки
        private void LoadPackersData()
        {
            packers = db.Packers.ToList();
           
            DataContext = this;
            cbPackers.DisplayMemberPath = "FIO";
            cbPackers.SelectedValuePath = "FIO";
            cbPackers.ItemsSource = db.Packers.ToList();

        }

        private void LoadFoundryData()
        {
            foundry = db.Foundry.ToList();

            DataContext = this;
            cbFoundry.DisplayMemberPath = "FIO_Foundry";
            cbFoundry.SelectedValuePath = "FIO_Foundry";
            cbFoundry.ItemsSource = db.Foundry.ToList();

        }

        private void cbPackers_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var packers = item as Packers;
                return packers.FIO.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void cbFoundry_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var foundry = item as Foundry;
                return foundry.FIO_Foundry.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void DopReport_Click(object sender, RoutedEventArgs e)
        {

            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите даты!");
                return;
            }

            GeneratePackerReportForOne(dpStartDate.SelectedDate.Value, dpEndDate.SelectedDate.Value);
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as MainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowFirstWindow();
            }
        }

        #endregion
        private void GeneratePackerReportForOne(DateTime startDate, DateTime endDate)
        {
            if (cbPackers.SelectedItem == null)
            {
                MessageBox.Show("Выберите упаковщицу!");
                return;
            }

            var selectedPacker = cbPackers.SelectedItem as Packers;
            string packerName = selectedPacker?.FIO;

            if (string.IsNullOrEmpty(packerName))
            {
                MessageBox.Show("Не удалось получить имя упаковщицы!");
                return;
            }

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
                    headerStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽";
                    moneyStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var summaryStyle = workbook.Style;
                    summaryStyle.Font.Bold = true;
                    summaryStyle.Fill.BackgroundColor = XLColor.Yellow;
                    summaryStyle.NumberFormat.NumberFormatId = 2;
                    summaryStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var defaultStyle = workbook.Style;
                    defaultStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    
                    worksheet.Cell(1, 1).Value = $"Отчет для упаковщицы {packerName} за период с {startDate:dd.MM.yyyy} по {endDate:dd.MM.yyyy}";
                    worksheet.Range(1, 1, 1, 18).Merge().Style = headerStyle;

                    // Берутся ненулевые артикулы
                    var usedArticles = db.DailyReport
                        .Where(dr => dr.FIO == packerName &&
                                     dr.DatePack >= startDate &&
                                     dr.DatePack <= endDate &&
                                     dr.Packs > 0)
                        .Select(dr => dr.ArticlePack)
                        .Distinct()
                        .OrderBy(a => a)
                        .ToList();

                    // для основной
                    string mainTableSql = @"
                    SELECT 
                        dr.DatePack,
                        dr.ArticlePack,
                        dr.Packs as Packs,
                        dr.FinePacks as Fines,
                        a.PricePackers as Price
                    FROM DailyReport dr
                    JOIN Articles a ON dr.ArticlePack = a.Article
                    WHERE dr.FIO = @packerName
                        AND dr.DatePack BETWEEN @startDate AND @endDate
                        AND dr.Packs > 0
                    ORDER BY dr.DatePack, dr.ArticlePack";

                    // вообще все для расчета штрафов и премий и тд
                    string allDataSql = @"
                    SELECT 
                        dr.DatePack,
                        dr.ArticlePack,
                        dr.Packs as Packs,
                        dr.FinePacks as Fines,
                        dr.FinePacksFoundry as FinePacksFoundry,
                        a.PricePackers as Price
                    FROM DailyReport dr
                    JOIN Articles a ON dr.ArticlePack = a.Article
                    WHERE dr.FIO = @packerName
                      AND dr.DatePack BETWEEN @startDate AND @endDate
                    ORDER BY dr.DatePack, dr.ArticlePack";

                    // заполнение основной табл
                    var mainTableData = new List<PackerReportItem>();
                    var allReportData = new List<PackerReportItem>();

                    using (var connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        // данные для основной
                        using (var cmd = new SqlCommand(mainTableSql, connection))
                        {
                            cmd.Parameters.AddWithValue("@packerName", packerName);
                            cmd.Parameters.AddWithValue("@startDate", startDate);
                            cmd.Parameters.AddWithValue("@endDate", endDate);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var item = new PackerReportItem
                                    {
                                        DatePack = (DateTime)reader["DatePack"],
                                        Article = reader["ArticlePack"].ToString(),
                                        Packs = Math.Round(Convert.ToDecimal(reader["Packs"]), 2),
                                        Fines = Convert.ToDecimal(reader["Fines"]),
                                        Price = Convert.ToDecimal(reader["Price"])
                                    };
                                    mainTableData.Add(item);
                                }
                            }
                        }

                        // данные для расчетов
                        using (var cmd = new SqlCommand(allDataSql, connection))
                        {
                            cmd.Parameters.AddWithValue("@packerName", packerName);
                            cmd.Parameters.AddWithValue("@startDate", startDate);
                            cmd.Parameters.AddWithValue("@endDate", endDate);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var item = new PackerReportItem
                                    {
                                        DatePack = (DateTime)reader["DatePack"],
                                        Article = reader["ArticlePack"].ToString(),
                                        Packs = Math.Round(Convert.ToDecimal(reader["Packs"]), 2),
                                        Fines = Convert.ToDecimal(reader["Fines"]),
                                        FinePacksFoundry = Convert.ToDecimal(reader["FinePacksFoundry"]),
                                        Price = Convert.ToDecimal(reader["Price"])
                                    };
                                    allReportData.Add(item);
                                }
                            }
                        }
                    }

                    // допы
                    var additionalServices = db.DopPackers
                        .Where(x => x.FIO == packerName && x.DateDopPackers >= startDate && x.DateDopPackers <= endDate)
                        .OrderBy(x => x.DateDopPackers)
                        .ToList();

                    decimal totalAdditionalServices = (decimal)additionalServices.Sum(x => x.Colvo * x.PriceForOne);

                    // расчеты
                    decimal totalPacks = allReportData.Sum(x => x.Packs);
                    decimal totalFines = allReportData.Sum(x => x.Fines);
                    decimal totalFinePacksFoundry = allReportData.Sum(x => x.FinePacksFoundry);
                    decimal totalSalary = mainTableData.Sum(x => x.Packs * x.Price);
                    decimal totalPremium = totalFinePacksFoundry * 2;
                    decimal fineAmount = totalFines;
                    

                    // заголовки
                    int row = 3;
                    int col = 1;

                    worksheet.Cell(row, col++).Value = "Дата";
                    foreach (var article in usedArticles)
                    {
                        worksheet.Cell(row, col++).Value = article;
                    }
                    worksheet.Range(row, 1, row, col - 1).Style = headerStyle;
                    row++;

                    // строка с ценами
                    col = 1;
                    worksheet.Cell(row, col++).Value = "Цена";
                    foreach (var article in usedArticles)
                    {
                        var price = db.Articles
                            .Where(x => x.Article == article)
                            .Select(x => x.PricePackers)
                            .FirstOrDefault();

                        worksheet.Cell(row, col).Value = price;
                        worksheet.Cell(row, col).Style = moneyStyle;
                        col++;
                    }
                    row++;

                    // Заполнение табл
                    foreach (var item in mainTableData.OrderBy(x => x.DatePack).ThenBy(x => x.Article))
                    {
                        col = 1;
                        worksheet.Cell(row, col++).Value = item.DatePack.ToString("dd.MM.yyyy");

                        foreach (var article in usedArticles)
                        {
                            decimal packs = item.Article == article ? item.Packs : 0m;
                            worksheet.Cell(row, col).Value = packs;
                            worksheet.Cell(row, col).Style.NumberFormat.NumberFormatId = 2;
                            col++;
                        }
                        row++;
                    }

                    // итог
                    int totalRow = row++;
                    int sumRow = row++;
                    int emptyRowAfterSum = row++; // пустая строчка после суммы
                    int packingRow = row++;
                    int fineRow = row++;
                    int premiumRow = row++;
                    int additionalServicesRow = row++;
                    int salaryRow = row++;

                    worksheet.Cell(totalRow, 1).Value = "Количество за месяц";
                    col = 2;

                    Dictionary<string, decimal> articleTotals = new Dictionary<string, decimal>();
                    foreach (var article in usedArticles)
                    {
                        decimal total = mainTableData
                            .Where(x => x.Article == article)
                            .Sum(x => x.Packs);

                        articleTotals[article] = total;
                        worksheet.Cell(totalRow, col).Value = total;
                        worksheet.Cell(totalRow, col).Style = summaryStyle;
                        col++;
                    }
                    worksheet.Range(totalRow, 1, totalRow, usedArticles.Count + 1).Style = summaryStyle;

                    // стр сумма
                    worksheet.Cell(sumRow, 1).Value = "Сумма";
                    col = 2;

                    foreach (var article in usedArticles)
                    {
                        decimal price = mainTableData
                            .Where(x => x.Article == article)
                            .Select(x => x.Price)
                            .FirstOrDefault();

                        decimal sum = articleTotals[article] * price;
                        worksheet.Cell(sumRow, col).Value = sum;
                        worksheet.Cell(sumRow, col).Style = summaryStyle;
                        worksheet.Cell(sumRow, col).Style.NumberFormat.Format = "#,##0.00 ₽";
                        col++;
                    }
                    worksheet.Range(sumRow, 1, sumRow, usedArticles.Count + 1).Style = summaryStyle;

                    // пустая строчка(?)
                    worksheet.Row(emptyRowAfterSum).Height = 20;

                    
                    worksheet.Cell(packingRow, 1).Value = "Упаковка:";
                    worksheet.Cell(packingRow, 2).Value = totalSalary;
                    worksheet.Cell(packingRow, 2).Style.NumberFormat.Format = "#,##0.00 ₽";

                    
                    worksheet.Cell(fineRow, 1).Value = "Штрафные пачки";
                    worksheet.Cell(fineRow, 2).Value = fineAmount;
                    

                    
                    worksheet.Cell(premiumRow, 1).Value = "Доп. плата за брак";
                    worksheet.Cell(premiumRow, 2).Value = totalPremium;
                    worksheet.Cell(premiumRow, 2).Style.NumberFormat.Format = "#,##0.00 ₽";

                    
                    worksheet.Cell(additionalServicesRow, 1).Value = "Доп. услуги";
                    worksheet.Cell(additionalServicesRow, 2).Value = totalAdditionalServices;
                    worksheet.Cell(additionalServicesRow, 2).Style.NumberFormat.Format = "#,##0.00 ₽";

                    
                    worksheet.Cell(salaryRow, 1).Value = "Итого";
                    worksheet.Cell(salaryRow, 2).Value = totalSalary + totalPremium  + totalAdditionalServices - totalFines*50;
                    worksheet.Cell(salaryRow, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                    worksheet.Cell(salaryRow, 2).Style.Font.Bold = true;

                    // таблица с допами
                    if (additionalServices.Any())
                    {
                        row = salaryRow + 2;

                        worksheet.Cell(row, 1).Value = "Дополнительные услуги упаковщицы";
                        worksheet.Range(row, 1, row, 2).Merge().Style = headerStyle;
                        row++;

                        worksheet.Cell(row, 1).Value = "Дата услуги";
                        worksheet.Cell(row, 2).Value = "Сумма";
                        worksheet.Range(row, 1, row, 2).Style = headerStyle;
                        row++;

                        foreach (var service in additionalServices)
                        {
                            worksheet.Cell(row, 1).Value = service.DateDopPackers.ToString("dd.MM.yyyy");
                            worksheet.Cell(row, 2).Value = service.Colvo * service.PriceForOne;
                            worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                            row++;
                        }

                        worksheet.Cell(row, 1).Value = "Итого:";
                        worksheet.Cell(row, 2).Value = totalAdditionalServices;
                        worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                        worksheet.Range(row, 1, row, 2).Style = summaryStyle;
                    }

                    // пустая стр после допов если есть
                    if (additionalServices.Any())
                    {
                        row++;
                        worksheet.Row(row).Height = 20;
                    }

                    // каждая доплата за брак в таблице
                    var finePacksData = allReportData
                        .Where(x => x.FinePacksFoundry > 0)
                        .OrderBy(x => x.DatePack)
                        .ToList();

                    if (finePacksData.Any())
                    {
                        row++;

                        worksheet.Cell(row, 1).Value = "Премии за брак (FinePacksFoundry × 2)";
                        worksheet.Range(row, 1, row, 2).Merge().Style = headerStyle;
                        row++;

                        worksheet.Cell(row, 1).Value = "Дата";
                        worksheet.Cell(row, 2).Value = "Сумма премии";
                        worksheet.Range(row, 1, row, 2).Style = headerStyle;
                        row++;

                        foreach (var item in finePacksData)
                        {
                            worksheet.Cell(row, 1).Value = item.DatePack.ToString("dd.MM.yyyy");
                            worksheet.Cell(row, 2).Value = item.FinePacksFoundry * 2;
                            worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                            row++;
                        }

                        worksheet.Cell(row, 1).Value = "Итого:";
                        worksheet.Cell(row, 2).Value = totalPremium;
                        worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                        worksheet.Range(row, 1, row, 2).Style = summaryStyle;
                    }

                    
                    worksheet.Columns().AdjustToContents();

                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_упаковщица_{packerName}_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

                    workbook.SaveAs(tempFilePath);
                    Process.Start(new ProcessStartInfo(tempFilePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        
        public class PackerReportItem
        {
            public DateTime DatePack { get; set; }
            public string Article { get; set; }
            public decimal Packs { get; set; }
            public decimal Fines { get; set; }
            public decimal FinePacksFoundry { get; set; }
            public decimal Price { get; set; }
        }

        private void DopReportFoundry_Click(object sender, RoutedEventArgs e)
        {
            if (dpStartDateFoundry.SelectedDate == null || dpEndDateFoundry.SelectedDate == null)
            {
                MessageBox.Show("Выберите даты!");
                return;
            }

            GenerateFoundryReportForOne(dpStartDateFoundry.SelectedDate.Value, dpEndDateFoundry.SelectedDate.Value);
        }


        private void GenerateFoundryReportForOne(DateTime startDate, DateTime endDate)
        {
            if (cbFoundry.SelectedItem == null)
            {
                MessageBox.Show("Выберите литейщика!");
                return;
            }

            var selectedFoundry = cbFoundry.SelectedItem as Foundry;
            string foundryName = selectedFoundry?.FIO_Foundry;

            if (string.IsNullOrEmpty(foundryName))
            {
                MessageBox.Show("Не удалось получить имя литейщика!");
                return;
            }

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
                    headerStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var moneyStyle = workbook.Style;
                    moneyStyle.NumberFormat.Format = "#,##0.00 ₽";
                    moneyStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var summaryStyle = workbook.Style;
                    summaryStyle.Font.Bold = true;
                    summaryStyle.Fill.BackgroundColor = XLColor.Yellow;
                    summaryStyle.NumberFormat.NumberFormatId = 2;
                    summaryStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    var defaultStyle = workbook.Style;
                    defaultStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(1, 1).Value = $"Отчет для литейщика {foundryName} за период с {startDate:dd.MM.yyyy} по {endDate:dd.MM.yyyy}";
                    worksheet.Range(1, 1, 1, 18).Merge().Style = headerStyle;

                    // данные для основной
                    var usedArticles = db.DailyReport
                        .Where(dr => dr.FIO_Foundry == foundryName &&
                                     dr.DatePack >= startDate &&
                                     dr.DatePack <= endDate &&
                                     dr.Packs2 > 0)
                        .Select(dr => dr.ArticleFoundry)
                        .Distinct()
                        .OrderBy(a => a)
                        .ToList();

                    // для основной
                    string mainTableSql = @"
                                            SELECT 
                                                dr.DateFoundry,
                                                dr.DatePack,
                                                dr.ArticleFoundry,
                                                dr.Packs2 as Packs,
                                                a.PriceFoundry
                                            FROM DailyReport dr
                                            JOIN Articles a ON dr.ArticleFoundry = a.Article
                                            WHERE dr.FIO_Foundry = @foundryName
                                              AND dr.DatePack BETWEEN @startDate AND @endDate
                                              AND dr.Packs2 > 0
                                            ORDER BY dr.DateFoundry, dr.DatePack, dr.ArticleFoundry";

                    // для расчетов
                    string allDataSql = @"
                                            SELECT 
                                                dr.DateFoundry,
                                                dr.DatePack,
                                                dr.ArticleFoundry,
                                                dr.Packs2 as Packs,
                                                dr.FinePacksFoundry as Fines,
                                                a.PriceFoundry
                                            FROM DailyReport dr
                                            JOIN Articles a ON dr.ArticleFoundry = a.Article
                                            WHERE dr.FIO_Foundry = @foundryName
                                              AND dr.DatePack BETWEEN @startDate AND @endDate
                                            ORDER BY dr.DateFoundry, dr.DatePack, dr.ArticleFoundry";

                    // заполнение основнойц
                    var mainTableData = new List<FoundryReportItem>();
                    var allReportData = new List<FoundryReportItem>();

                    using (var connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        // данные для основной
                        using (var cmd = new SqlCommand(mainTableSql, connection))
                        {
                            cmd.Parameters.AddWithValue("@foundryName", foundryName);
                            cmd.Parameters.AddWithValue("@startDate", startDate);
                            cmd.Parameters.AddWithValue("@endDate", endDate);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var item = new FoundryReportItem
                                    {
                                        DateFoundry = (DateTime)reader["DateFoundry"],
                                        DatePack = (DateTime)reader["DatePack"],
                                        Article = reader["ArticleFoundry"].ToString(),
                                        Packs = Math.Round(Convert.ToDecimal(reader["Packs"]), 2),
                                        Price = Convert.ToDecimal(reader["PriceFoundry"])
                                    };
                                    mainTableData.Add(item);
                                }
                            }
                        }

                        // для расчетов
                        using (var cmd = new SqlCommand(allDataSql, connection))
                        {
                            cmd.Parameters.AddWithValue("@foundryName", foundryName);
                            cmd.Parameters.AddWithValue("@startDate", startDate);
                            cmd.Parameters.AddWithValue("@endDate", endDate);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var item = new FoundryReportItem
                                    {
                                        DateFoundry = (DateTime)reader["DateFoundry"],
                                        DatePack = (DateTime)reader["DatePack"],
                                        Article = reader["ArticleFoundry"].ToString(),
                                        Packs = Math.Round(Convert.ToDecimal(reader["Packs"]), 2),
                                        Fines = Convert.ToDecimal(reader["Fines"]),
                                        Price = Convert.ToDecimal(reader["PriceFoundry"])
                                    };
                                    allReportData.Add(item);
                                }
                            }
                        }
                    }

                    // допы
                    var additionalServices = db.DopFoundry
                        .Where(x => x.FIO_Foundry == foundryName && x.DateDop >= startDate && x.DateDop <= endDate)
                        .GroupBy(x => x.DateDop)
                        .Select(g => new
                        {
                            Date = g.Key,
                            Total = g.Sum(x => x.Colvo * x.PriceForOne)
                        })
                        .OrderBy(x => x.Date)
                        .ToList();

                    decimal totalAdditionalServices = additionalServices.Sum(x => x.Total);

                    
                    decimal totalPacks = allReportData.Sum(x => x.Packs);
                    decimal totalFines = allReportData.Sum(x => x.Fines);
                    decimal totalSalary = mainTableData.Sum(x => x.Packs * x.Price); 
                    decimal fineAmount = 0m;

                    if (totalPacks > 0 && totalFines > 0)
                    {
                        decimal allowedFines = totalPacks * 0.05m;
                        if (totalFines > allowedFines)
                        {
                            fineAmount = (totalFines - allowedFines) * 12m;
                        }
                    }

                    // заголовки
                    int row = 3;
                    int col = 2; // с 2 колонки, тк 1 - "Дата сборки"

                    worksheet.Cell(row, 1).Value = "Дата сборки";
                    worksheet.Cell(row, col++).Value = "Дата литья";
                    foreach (var article in usedArticles)
                    {
                        worksheet.Cell(row, col++).Value = article;
                    }
                    worksheet.Range(row, 1, row, col - 1).Style = headerStyle;
                    row++;

                    // строка с ценами
                    col = 2;
                    worksheet.Cell(row, 1).Value = "Цена";
                    worksheet.Cell(row, col++).Value = ""; // пусто для даты литья
                    foreach (var article in usedArticles)
                    {
                        var price = db.Articles
                            .Where(x => x.Article == article)
                            .Select(x => x.PriceFoundry)
                            .FirstOrDefault();

                        worksheet.Cell(row, col).Value = price;
                        worksheet.Cell(row, col).Style = moneyStyle;
                        col++;
                    }
                    row++;

                    // заполнение
                    foreach (var item in mainTableData.OrderBy(x => x.DatePack).ThenBy(x => x.DateFoundry).ThenBy(x => x.Article))
                    {
                        col = 2;
                        worksheet.Cell(row, 1).Value = item.DatePack.ToString("dd.MM.yyyy");
                        worksheet.Cell(row, col++).Value = item.DateFoundry.ToString("dd.MM.yyyy");

                        foreach (var article in usedArticles)
                        {
                            decimal packs = item.Article == article ? item.Packs : 0m;
                            worksheet.Cell(row, col).Value = packs;
                            worksheet.Cell(row, col).Style.NumberFormat.NumberFormatId = 2;
                            col++;
                        }
                        row++;
                    }

                    // итог строчки
                    int totalRow = row++;
                    int sumRow = row++;
                    int emptyRowAfterSum = row++; // пустая после суммы
                    int foundryRow = row++; 
                    int fineRow = row++;
                    int additionalServicesRow = row++;
                    int salaryRow = row++;

                    worksheet.Cell(totalRow, 1).Value = "Количество за месяц";
                    worksheet.Cell(totalRow, 2).Value = ""; 
                    col = 3;

                    Dictionary<string, decimal> articleTotals = new Dictionary<string, decimal>();
                    foreach (var article in usedArticles)
                    {
                        decimal total = mainTableData
                            .Where(x => x.Article == article)
                            .Sum(x => x.Packs);

                        articleTotals[article] = total;
                        worksheet.Cell(totalRow, col).Value = total;
                        worksheet.Cell(totalRow, col).Style = summaryStyle;
                        col++;
                    }
                    worksheet.Range(totalRow, 1, totalRow, usedArticles.Count + 2).Style = summaryStyle;

                    worksheet.Cell(sumRow, 1).Value = "Сумма";
                    worksheet.Cell(sumRow, 2).Value = "";
                    col = 3;

                    foreach (var article in usedArticles)
                    {
                        decimal price = mainTableData
                            .Where(x => x.Article == article)
                            .Select(x => x.Price)
                            .FirstOrDefault();

                        decimal sum = articleTotals[article] * price;
                        worksheet.Cell(sumRow, col).Value = sum;
                        worksheet.Cell(sumRow, col).Style = summaryStyle;
                        worksheet.Cell(sumRow, col).Style.NumberFormat.Format = "#,##0.00 ₽";
                        col++;
                    }
                    worksheet.Range(sumRow, 1, sumRow, usedArticles.Count + 2).Style = summaryStyle;

                    // пустая строка после суммы
                    worksheet.Row(emptyRowAfterSum).Height = 20;

                    
                    worksheet.Cell(foundryRow, 1).Value = "Литье:";
                    worksheet.Cell(foundryRow, 3).Value = totalSalary;
                    worksheet.Cell(foundryRow, 3).Style.NumberFormat.Format = "#,##0.00 ₽";

                    
                    worksheet.Cell(fineRow, 1).Value = "Штраф";
                    worksheet.Cell(fineRow, 2).Value = ""; 
                    worksheet.Cell(fineRow, 3).Value = fineAmount;
                    worksheet.Cell(fineRow, 3).Style.NumberFormat.Format = "#,##0.00 ₽";

                    
                    worksheet.Cell(additionalServicesRow, 1).Value = "Дополнительные услуги";
                    worksheet.Cell(additionalServicesRow, 2).Value = ""; 
                    worksheet.Cell(additionalServicesRow, 3).Value = totalAdditionalServices;
                    worksheet.Cell(additionalServicesRow, 3).Style.NumberFormat.Format = "#,##0.00 ₽";

                    
                    worksheet.Cell(salaryRow, 1).Value = "Итого";
                    worksheet.Cell(salaryRow, 2).Value = ""; 
                    worksheet.Cell(salaryRow, 3).Value = totalSalary - fineAmount + totalAdditionalServices;
                    worksheet.Cell(salaryRow, 3).Style.NumberFormat.Format = "#,##0.00 ₽";
                    worksheet.Cell(salaryRow, 3).Style.Font.Bold = true;

                    // ддддопы
                    if (additionalServices.Any())
                    {
                        row = salaryRow + 2; 

                        
                        worksheet.Cell(row, 1).Value = "Дополнительные услуги";
                        worksheet.Range(row, 1, row, 2).Merge().Style = headerStyle;
                        row++;

                        
                        worksheet.Cell(row, 1).Value = "Дата услуги";
                        worksheet.Cell(row, 2).Value = "Сумма";
                        worksheet.Range(row, 1, row, 2).Style = headerStyle;
                        row++;

                        foreach (var service in additionalServices)
                        {
                            worksheet.Cell(row, 1).Value = service.Date.ToString("dd.MM.yyyy");
                            worksheet.Cell(row, 2).Value = service.Total;
                            worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                            row++;
                        }

                        
                        worksheet.Cell(row, 1).Value = "Итого:";
                        worksheet.Cell(row, 2).Value = totalAdditionalServices;
                        worksheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00 ₽";
                        worksheet.Range(row, 1, row, 2).Style = summaryStyle;
                    }

                    
                    worksheet.Columns().AdjustToContents();

                    
                    var tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"Отчет_литейщик_{foundryName}_{startDate:dd.MM.yyyy}_по_{endDate:dd.MM.yyyy}.xlsx");

                    workbook.SaveAs(tempFilePath);
                    Process.Start(new ProcessStartInfo(tempFilePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        
        public class FoundryReportItem
        {
            public DateTime DateFoundry { get; set; }
            public DateTime DatePack { get; set; }
            public string Article { get; set; }
            public decimal Packs { get; set; }
            public decimal Fines { get; set; }
            public decimal Price { get; set; }
        }
    }
}
