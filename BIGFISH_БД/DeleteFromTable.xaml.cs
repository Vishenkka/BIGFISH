using System;
using System.Collections.Generic;
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

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для DeleteFromTable.xaml
    /// </summary>
    public partial class DeleteFromTable : Window
    {
        public DeleteFromTable()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите обе даты!");
                return;
            }

            DateTime startDate = dpStartDate.SelectedDate.Value.Date;
            DateTime endDate = dpEndDate.SelectedDate.Value.Date.AddDays(1).AddSeconds(-1); // До конца дня

            var confirmResult = MessageBox.Show(
                $"Вы действительно хотите удалить все записи с {startDate:dd.MM.yyyy} по {endDate:dd.MM.yyyy}?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmResult != MessageBoxResult.Yes)
            {
                return;
            }

            string connectionString = @"data source=V_ISHENKA\SQLEXPRESS;initial catalog=BigFishBD;integrated security=True;";

            try
            {
                using (var db = new BigFishBDEntities())
                {
                    var recordsToDelete = db.DailyReport
                        .Where(dr => dr.DatePack >= startDate && dr.DatePack <= endDate)
                        .ToList();

                    if (!recordsToDelete.Any())
                    {
                        MessageBox.Show("Нет записей для удаления в выбранном периоде.");
                        return;
                    }

                    db.DailyReport.RemoveRange(recordsToDelete);
                    int affectedRows = db.SaveChanges();

                    MessageBox.Show($"Успешно удалено {affectedRows} записей.",
                                  "Удаление завершено",
                                  MessageBoxButton.OK,
                                  MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении данных:\n{ex.Message}",
                              "Ошибка",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
            }
        }
    }
}
