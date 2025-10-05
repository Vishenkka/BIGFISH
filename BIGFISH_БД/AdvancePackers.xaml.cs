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
    /// Логика взаимодействия для AdvancePackers.xaml
    /// </summary>
    public partial class AdvancePackers : UserControl
    {
        public AdvancePackers()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
            LoadPackersData();
        }

        BigFishBDEntities db;
        public List<AdvancePayPackers> advancePayPackers { get; set; }
        public List<Packers> packers { get; set; }
       

        private void LoadPackersData()
        {
            packers = db.Packers.ToList();
            tbPackers.ItemsSource = packers;
            DataContext = this;

        }
        
        private void combobind()
        {
            TablAdvance.ItemsSource = db.AdvancePayPackers.ToList();
            DataContext = this;
        }
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablAdvance.ItemsSource = db.AdvancePayPackers.ToList();
        }

        private void tbPackers_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var packers = item as Packers;
                return packers.FIO.ToLower().Contains(combo.Text.ToLower());
            };
        }


        private void Add_Click(object sender, RoutedEventArgs e)
        {


            try
            {
                DateTime DateAdv = dpDateAdv.SelectedDate.Value;

                AdvancePayPackers cl = new AdvancePayPackers();
                cl.DateAdv = DateAdv;
                cl.FIO = tbPackers.Text;

                cl.AdvancePay = (float)Math.Round(float.Parse(tbSum.Text), 2);

                db.AdvancePayPackers.Add(cl);
                db.SaveChanges();
                TablAdvance.ItemsSource = db.AdvancePayPackers.ToList();


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
                var SelectDN = db.AdvancePayPackers.Where(m => m.Id == SDN).FirstOrDefault();
                db.AdvancePayPackers.Remove(SelectDN);
                db.SaveChanges();
                TablAdvance.ItemsSource = db.AdvancePayPackers.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись");
            }
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
