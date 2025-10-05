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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для AdminFirstWindow.xaml
    /// </summary>
    public partial class AdminFirstWindow : UserControl
    {
        public AdminFirstWindow()
        {
            InitializeComponent();
        }

        private void DailyReport_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as AdminMainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowDailyReport();
            }
        }

        private void ReportPackers_Click(object sender, RoutedEventArgs e)
        {
            var datePickPackers = new DatePickPackers();
            datePickPackers.ShowDialog();

        }

        private void ReportFoundry_Click(object sender, RoutedEventArgs e)
        {
            var datePickFoundry = new AdminMainWindow();
            datePickFoundry.ShowDialog();
        }

        private void TotalReport_Click(object sender, RoutedEventArgs e)
        {
            var datePickTotal = new DatePickTotal();
            datePickTotal.ShowDialog();
        }

        private void DopReport_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as AdminMainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowChooseDop();
            }
        }

        private void OnePersonReport_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as AdminMainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowChooseForOnePerson();
            }
        }
    }
}
