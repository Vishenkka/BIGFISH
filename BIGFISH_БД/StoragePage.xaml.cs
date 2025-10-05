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
    /// Логика взаимодействия для StoragePage.xaml
    /// </summary>
    public partial class StoragePage : UserControl
    {
        BigFishBDEntities db;
        public List<Storage> storage { get; set; }
        public StoragePage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();

        }

        private void combobind()
        {

            storage = db.Storage.ToList();
            cbSpisok.ItemsSource = storage;
            cbSpisok.DisplayMemberPath = "StorageName";
            DataContext = this;
        }

        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablStorage.ItemsSource = storage.Where(x => x.StorageName.Contains(tbPoisk.Text)).ToList();

        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Storage cl = new Storage();
                cl.StorageName = tbStorage.Text;
                db.Storage.Add(cl);
                db.SaveChanges();
                TablStorage.ItemsSource = db.Storage.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Заполните поле!");
            }

        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string SDN = tbStorage.Text;
                var SelectDN = db.Storage.Where(m => m.StorageName == SDN).FirstOrDefault();
                db.Storage.Remove(SelectDN);
                db.SaveChanges();
                TablStorage.ItemsSource = db.Storage.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Пустое поле! Заполните его вручную");
            }


        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablStorage.ItemsSource = db.Storage.ToList();
            cbSpisok.ItemsSource = db.Storage.ToList();
        }

        private void cbSpisok_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = (Storage)cbSpisok.SelectedItem;
            TablStorage.ItemsSource = storage.Where(x => x.StorageName == selected.StorageName).ToList();
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Window.GetWindow(this) as MainWindow;
            var adminMainWindow = Window.GetWindow(this) as AdminMainWindow;
            if (mainWindow != null)
            {
                mainWindow.ShowFirstWindow();
            }
            else
            {
                adminMainWindow.ShowAdminFirstWindow();
            }
        }
    }
}
