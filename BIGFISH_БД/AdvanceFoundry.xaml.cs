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
    /// Логика взаимодействия для AdvanceFoundry.xaml
    /// </summary>
    public partial class AdvanceFoundry : UserControl
    {
        public AdvanceFoundry()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
            LoadFoundryData();
        }

   

        BigFishBDEntities db;
        public List<AdvancePayFoundry> advancePayFoundry { get; set; }
        public List<Foundry> foundry { get; set; }


        private void LoadFoundryData()
        {
            foundry = db.Foundry.ToList();
            tbFoundry.ItemsSource = foundry;
            DataContext = this;

        }

        private void combobind()
        {
            TablAdvance.ItemsSource = db.AdvancePayFoundry.ToList();
            DataContext = this;
        }
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablAdvance.ItemsSource = db.AdvancePayFoundry.ToList();
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


        private void Add_Click(object sender, RoutedEventArgs e)
        {


            try
            {
                DateTime DateAdv = dpDateAdv.SelectedDate.Value;

                AdvancePayFoundry cl = new AdvancePayFoundry();
                cl.DateAdv = DateAdv;
                cl.FIO_Foundry = tbFoundry.Text;

                cl.AdvancePay = (float)Math.Round(float.Parse(tbSum.Text), 2);

                db.AdvancePayFoundry.Add(cl);
                db.SaveChanges();
                TablAdvance.ItemsSource = db.AdvancePayFoundry.ToList();


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
                var SelectDN = db.AdvancePayFoundry.Where(m => m.Id == SDN).FirstOrDefault();
                db.AdvancePayFoundry.Remove(SelectDN);
                db.SaveChanges();
                TablAdvance.ItemsSource = db.AdvancePayFoundry.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись");
            }
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
