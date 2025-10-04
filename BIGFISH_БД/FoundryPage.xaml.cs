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
    /// Логика взаимодействия для FoundryPage.xaml
    /// </summary>
    public partial class FoundryPage : UserControl
    {
        BigFishBDEntities db;
        public List<Foundry> foundry { get; set; }
        public FoundryPage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
        }

        private void combobind()
        {

            foundry = db.Foundry.ToList();
            cbSpisok.ItemsSource = foundry;
            cbSpisok.DisplayMemberPath = "FIO_Foundry";
            DataContext = this;
        }

        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablFoundry.ItemsSource = foundry.Where(x => x.FIO_Foundry.Contains(tbPoisk.Text)).ToList();

        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Foundry cl = new Foundry();
                cl.FIO_Foundry = tbFIO.Text;
                db.Foundry.Add(cl);
                db.SaveChanges();
                TablFoundry.ItemsSource = db.Foundry.ToList();
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
                var SelectDN = db.Foundry.Where(m => m.FIO_Foundry == SDN).FirstOrDefault();
                db.Foundry.Remove(SelectDN);
                db.SaveChanges();
                TablFoundry.ItemsSource = db.Foundry.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись.");
            }


        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablFoundry.ItemsSource = db.Foundry.ToList();
            cbSpisok.ItemsSource = db.Foundry.ToList();
        }

        private void cbSpisok_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedClient = (Foundry)cbSpisok.SelectedItem;
            TablFoundry.ItemsSource = foundry.Where(x => x.FIO_Foundry == selectedClient.FIO_Foundry).ToList();
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
