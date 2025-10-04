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
    /// Логика взаимодействия для DopInsert.xaml
    /// </summary>
    public partial class DopInsert : UserControl
    {
        BigFishBDEntities db;
        public List<DopFoundry> dopFoundry { get; set; }
        public List<Foundry> foundry { get; set; }
        public List<Naim_Dop> naim_Dop { get; set; }

        public DopInsert()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
            LoadFoundryData();
            LoadNaimDopData();
        }

        private void LoadFoundryData()
        {
            foundry = db.Foundry.ToList();
            tbFoundry.ItemsSource = foundry;
            DataContext = this;

        }
        private void LoadNaimDopData()
        {
            naim_Dop = db.Naim_Dop.ToList();
            tbNameDop.ItemsSource = naim_Dop;
            DataContext = this;
        }
        private void combobind()
        {
            dopFoundry = db.DopFoundry.ToList();
            DataContext = this;
        }
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablDopRep.ItemsSource = db.DopFoundry.ToList();
        }

        private void tbFoundry_TextChanged(object sender, TextChangedEventArgs e)
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

        private void tbNameDop_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var name_Dop = item as Naim_Dop;
                return name_Dop.Name_Dop.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {


            try
            {
                DateTime DateDop = dpDateDop.SelectedDate.Value;

                DopFoundry cl = new DopFoundry();
                cl.DateDop = DateDop;
                cl.FIO_Foundry = tbFoundry.Text;
                cl.IdVed = tbIdVed.Text;
                cl.Name_Dop = tbNameDop.Text;
                cl.Colvo = Convert.ToInt32(tbColvo.Text);
                cl.PriceForOne = Convert.ToInt32(tbPriceForOne.Text);

                db.DopFoundry.Add(cl);
                db.SaveChanges();
                TablDopRep.ItemsSource = db.DopFoundry.ToList();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось добавить запись.");
            }

        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int SDN = Convert.ToInt32(tbId.Text);
                var SelectDN = db.DopFoundry.Where(m => m.IdDop == SDN).FirstOrDefault();
                db.DopFoundry.Remove(SelectDN);
                db.SaveChanges();
                TablDopRep.ItemsSource = db.DopFoundry.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись");
            }
        }

        private void AddNumVed_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int StrId = Convert.ToInt32(tbId.Text);
                var SelectId = db.DopFoundry.Where(w => w.IdDop == StrId).FirstOrDefault();
                SelectId.IdVed = tbNumVedBottom.Text;

                db.SaveChanges();
                TablDopRep.ItemsSource = db.DopFoundry.ToList();
            }
            catch (Exception ex) { MessageBox.Show("Не удалось добавить номер ведомости"); }

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
