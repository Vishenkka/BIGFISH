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
    /// Логика взаимодействия для FilterPackers.xaml
    /// </summary>
    public partial class FilterPackers : Window
    {
        public DateTime? SelectedStartDate { get; private set; }
        public DateTime? SelectedEndDate { get; private set; }
        public Packers SelectedPacker { get; private set; }
        public bool IsFilterApplied { get; private set; }

        public FilterPackers()
        {
            InitializeComponent();
            LoadPackers();
        }

        private void LoadPackers()
        {
            try
            {
                using (var db = new BigFishBDEntities())
                {
                    PackerComboBox.ItemsSource = db.Packers.AsNoTracking().ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки упаковщиц: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateDates())
                return;

            SelectedStartDate = StartDatePicker.SelectedDate;
            SelectedEndDate = EndDatePicker.SelectedDate;
            SelectedPacker = PackerComboBox.SelectedItem as Packers;
            IsFilterApplied = true;

            DialogResult = true;
            Close();
        }

        private bool ValidateDates()
        {
            if (!StartDatePicker.SelectedDate.HasValue || !EndDatePicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Выберите обе даты", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (StartDatePicker.SelectedDate > EndDatePicker.SelectedDate)
            {
                MessageBox.Show("Дата начала не может быть позже даты конца", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
