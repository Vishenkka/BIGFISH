using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
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
    /// Логика взаимодействия для AdminMainWindow.xaml
    /// </summary>
    public partial class AdminMainWindow : Window
    {
        public AdminMainWindow()
        {
            InitializeComponent();
            ShowAdminFirstWindow();
        }

        

        private void ButtonOpenMenu_Click(object sender, RoutedEventArgs e)
        {
            if (ButtonOpenMenu.Visibility == Visibility.Visible)
            {
                ButtonOpenMenu.Visibility = Visibility.Collapsed;
                ButtonCloseMenu.Visibility = Visibility.Visible;
            }
            else if (ButtonOpenMenu.Visibility == Visibility.Collapsed)
            {
                ButtonOpenMenu.Visibility = Visibility.Visible;
                ButtonCloseMenu.Visibility = Visibility.Collapsed;
            }
            //    ButtonCloseMenu.Visibility = Visibility.Visible;
            //ButtonOpenMenu.Visibility = Visibility.Collapsed;
        }

        private void ListViewMenu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UserControl usc = null;
            GridMain.Children.Clear();

            switch (((ListViewItem)((ListView)sender).SelectedItem).Name)
            {
                

                case "ItemStatistics":
                    usc = new StatisticPage();
                    GridMain.Children.Add(usc);
                    break;

                case "ItemStatisticsArticle":
                    usc = new StatisticArticlePage();
                    GridMain.Children.Add(usc);
                    break;

                case "ItemStatisticsFoundry":
                    usc = new StatisticOnePerson();
                    GridMain.Children.Add(usc);
                    break;

                case "ItemStatisticsPackers":
                    usc = new StatisticOnePacker();
                    GridMain.Children.Add(usc);
                    break;


                default: break;
            }


        }
        public void ShowDailyReport()
        {
            GridMain.Children.Clear();
            var dailyReportPage = new DailyReportPage();
            GridMain.Children.Add(dailyReportPage);
        }

        public void ShowChooseForOnePerson()
        {
            GridMain.Children.Clear();
            var chooseForOne = new ChooseForOne();
            GridMain.Children.Add(chooseForOne);
        }

        public void ShowDopReport()
        {
            GridMain.Children.Clear();
            var dopInsert = new DopInsert();
            GridMain.Children.Add(dopInsert);
        }

        public void ShowAdvPackers()
        {
            GridMain.Children.Clear();
            var advancePackers = new AdvancePackers();
            GridMain.Children.Add(advancePackers);
        }

        public void ShowAdvFoundry()
        {
            GridMain.Children.Clear();
            var advanceFoundry = new AdvanceFoundry();
            GridMain.Children.Add(advanceFoundry);
        }

        public void ShowChooseDop()
        {
            GridMain.Children.Clear();
            var chooseDop = new ChooseDopFrame();
            GridMain.Children.Add(chooseDop);
        }

        public void ShowDopPackers()
        {
            GridMain.Children.Clear();
            var dopPackers1 = new DopPackers1();
            GridMain.Children.Add(dopPackers1);
        }
        public void ShowAdminFirstWindow()
        {
            GridMain.Children.Clear();
            var adminFirstWindow = new AdminFirstWindow();
            GridMain.Children.Add(adminFirstWindow);
        }

        private void Contact_Click(object sender, RoutedEventArgs e)
        {
            GridMain.Children.Clear();
            var contacts = new Contacts();
            GridMain.Children.Add(contacts);
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("-");
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("-");

        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {

            DeleteFromTable deleteFromTable = new DeleteFromTable();
            deleteFromTable.Show();
        }
    }
}
