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
    /// Логика взаимодействия для DopPage.xaml
    /// </summary>
    public partial class DopPage : UserControl
    {
        BigFishBDEntities db;
        public List<Naim_Dop> naim_Dop { get; set; }
        public DopPage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();

            combobind();
        }

        private void combobind()
        {

            naim_Dop = db.Naim_Dop.ToList();
            cbSpisok.ItemsSource = naim_Dop;
            cbSpisok.DisplayMemberPath = "Name_Dop";
            DataContext = this;
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablDop.ItemsSource = db.Naim_Dop.ToList();
            cbSpisok.ItemsSource = db.Naim_Dop.ToList();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Naim_Dop cl = new Naim_Dop();
                cl.Name_Dop = tbNameDop.Text;
                db.Naim_Dop.Add(cl);
                db.SaveChanges();
                TablDop.ItemsSource = db.Naim_Dop.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не заполнено одно из полей.");
            }

        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string SDN = tbNameDop.Text;
                var SelectDN = db.Naim_Dop.Where(m => m.Name_Dop == SDN).FirstOrDefault();
                db.Naim_Dop.Remove(SelectDN);
                db.SaveChanges();
                TablDop.ItemsSource = db.Naim_Dop.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись (возможно, в таблице с записями допов есть запись с этим наименованием)");
            }


        }
        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablDop.ItemsSource = naim_Dop.Where(x => x.Name_Dop.Contains(tbPoisk.Text)).ToList();


        }

        private void cbSpisok_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = (Naim_Dop)cbSpisok.SelectedItem;
            TablDop.ItemsSource = naim_Dop.Where(x => x.Name_Dop == selected.Name_Dop).ToList();
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
