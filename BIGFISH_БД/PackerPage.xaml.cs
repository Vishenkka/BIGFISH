
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для PackerPage.xaml
    /// </summary>
    public partial class PackerPage : UserControl
    {
        BigFishBDEntities db;
        public List<Packers> packers { get; set; }



        public PackerPage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
        }

        private void combobind()
        {

            packers = db.Packers.ToList();
            cbSpisok.ItemsSource = packers;
            cbSpisok.DisplayMemberPath = "FIO";
            DataContext = this;
        }

        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablPackers.ItemsSource = packers.Where(x => x.FIO.Contains(tbPoisk.Text)).ToList();

        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Packers cl = new Packers();
                cl.FIO = tbFIO.Text;
                db.Packers.Add(cl);
                db.SaveChanges();
                TablPackers.ItemsSource = db.Packers.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось добавить запись.");
            }

        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string SDN = tbFIO.Text;
                var SelectDN = db.Packers.Where(m => m.FIO == SDN).FirstOrDefault();
                db.Packers.Remove(SelectDN);
                db.SaveChanges();
                TablPackers.ItemsSource = db.Packers.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись.");
            }


        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablPackers.ItemsSource = db.Packers.ToList();
            cbSpisok.ItemsSource = db.Packers.ToList();
        }

        private void cbSpisok_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedClient = (Packers)cbSpisok.SelectedItem;
            TablPackers.ItemsSource = packers.Where(x => x.FIO == selectedClient.FIO).ToList();
        }
        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as MainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowFirstWindow();
            }
        }

    }
}
