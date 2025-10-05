using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Runtime.InteropServices;
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
using static MaterialDesignThemes.Wpf.Theme;

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для DailyReportPage.xaml
    /// </summary>
    public partial class DailyReportPage : UserControl
    {
        BigFishBDEntities db;
        public List<DailyReport> dailyReport { get; set; }
        public List<Articles> articles { get; set; }
        public List<Packers> packers { get; set; }
        public List<Foundry> foundry { get; set; }
        public List<Storage>  storageName { get; set; }

        private List<Control> focusOrder = new List<Control>();

        public DailyReportPage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
            LoadArticlesData();
            LoadPackersData();
            LoadStorageData();
            LoadFoundryData();
            LoadColorArticles2Data();
            InitializeFocusOrder();
            this.Loaded += (s, e) => { focusOrder[0]?.Focus(); };



        }

        #region переключение по энтеру (мб переделать?)
        private void InitializeFocusOrder()
        {
            focusOrder = new List<Control>
            {
                tbPoisk, tbPoisk1, tbNumVedPoisk,
                dpDatePack, tbPacker, tbArticle, 
                tbStorageName, tbFinePacks, tbcolvoPacks,
                dpDateFoundry, tbFoundry, tbArticle2,
                tbcolvoPacks2, tbFinePacksFoundry,
                tbId, tbNumVed
            };
        }

        private void UIElement_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                if (sender is Control currentControl)
                {
                    // Находим индекс текущего элемента
                    int currentIndex = focusOrder.IndexOf(currentControl);

                    if (currentIndex >= 0 && currentIndex < focusOrder.Count - 1)
                    {
                        // Переходим к следующему элементу
                        focusOrder[currentIndex + 1].Focus();

                        // Для TextBox выделяем весь текст
                        if (focusOrder[currentIndex + 1] is System.Windows.Controls.TextBox nextTextBox)
                            nextTextBox.SelectAll();
                    }
                    else
                    {
                        // Если достигли конца - переходим к первому элементу
                        focusOrder[0].Focus();
                        if (focusOrder[0] is System.Windows.Controls.TextBox firstTextBox)
                            firstTextBox.SelectAll();
                    }
                }
            }
        }

        #endregion


        #region загрузки и TextChanged

        private void LoadArticlesData()
        {
            articles = db.Articles.ToList();
            tbArticle.ItemsSource = articles;
            DataContext = this;

        }
        private void LoadColorArticles2Data()
        {
            articles = db.Articles.ToList();
            tbArticle2.ItemsSource = articles;
            DataContext = this;

        }
        private void LoadFoundryData()
        {
            foundry = db.Foundry.ToList();
            tbFoundry.ItemsSource = foundry;
            DataContext = this;

        }
        private void LoadPackersData()
        {
            packers = db.Packers.ToList();
            tbPacker.ItemsSource = packers;
            DataContext = this;

        }
        private void LoadStorageData()
        {
            storageName = db.Storage.ToList();
            tbStorageName.ItemsSource = storageName;
            DataContext = this;


        }
        private void combobind()
        {

            dailyReport = db.DailyReport.ToList();
            DataContext = this;
        }
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablDailyRep.ItemsSource = db.DailyReport.ToList();
            if (TablDailyRep.Items.Count > 0)
            {
                var lastItem = TablDailyRep.Items[TablDailyRep.Items.Count - 1];
                TablDailyRep.ScrollIntoView(lastItem);

                
            }

        }

        private void tbFoundry_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as System.Windows.Controls.ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var foundry = item as Foundry;
                return foundry.FIO_Foundry.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void tbPacker_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as System.Windows.Controls.ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var packers = item as Packers;
                return packers.FIO.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void tbArticle_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as System.Windows.Controls.ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var article = item as Articles;
                return article.Article.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void tbArticle2_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as System.Windows.Controls.ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var article = item as Articles;
                return article.Article.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void tbStorageName_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as System.Windows.Controls.ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var storageName = item as Storage;
                return storageName.StorageName.ToLower().Contains(combo.Text.ToLower());
            };
        }

        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablDailyRep.ItemsSource = dailyReport.Where(x => x.FIO.Contains(tbPoisk.Text)).ToList();

        }

        private void tbPoisk1_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablDailyRep.ItemsSource = dailyReport.Where(x => x.FIO_Foundry.Contains(tbPoisk1.Text)).ToList();

        }

        private void tbNumVedPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablDailyRep.ItemsSource = dailyReport.Where(x => x.Number_Ved == tbNumVedPoisk.Text).ToList();

        }


        #endregion

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

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            

            try
            {
                DailyReport cl = new DailyReport();
                if (dpDatePack != null) 
                {
                    DateTime DatePacks = dpDatePack.SelectedDate.Value;
                    cl.DatePack = DatePacks;
                } ;

                if (dpDateFoundry != null)
                {
                    DateTime DateFoundry = dpDateFoundry.SelectedDate.Value;
                    cl.DateFoundry = DateFoundry;

                }
                ;


                   
                cl.FIO = tbPacker.Text;
                cl.ArticlePack = tbArticle.Text;
                cl.StorageName = tbStorageName.Text;

                cl.FinePacks = float.Parse(tbFinePacks.Text);

                cl.Packs = (float)Math.Round(float.Parse(tbcolvoPacks.Text),2);
                cl.ArticleFoundry = tbArticle2.Text;
                cl.Packs2 = (float)Math.Round(float.Parse(tbcolvoPacks2.Text),2);
                cl.FinePacksFoundry = (float)Math.Round(float.Parse(tbFinePacksFoundry.Text),2);
                cl.FIO_Foundry = tbFoundry.Text;

                cl.Number_Ved = tbNumVed.Text;

                db.DailyReport.Add(cl);
                db.SaveChanges();
                TablDailyRep.ItemsSource = db.DailyReport.ToList();

                
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
                var SelectDN = db.DailyReport.Where(m => m.Id == SDN).FirstOrDefault();
                db.DailyReport.Remove(SelectDN);
                db.SaveChanges();
                TablDailyRep.ItemsSource = db.DailyReport.ToList();
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
                var SelectId = db.DailyReport.Where(w => w.Id == StrId).FirstOrDefault();
                SelectId.Number_Ved = tbNumVed.Text;

                db.SaveChanges();
                TablDailyRep.ItemsSource = db.DailyReport.ToList();
            }
            catch (Exception ex) { MessageBox.Show("Не удалось добавить номер ведомости"); }

        }

        private void ShowChanges_Click(object sender, RoutedEventArgs e)
        {
            ChangeDailyReport changeDailyReport = new ChangeDailyReport();
            changeDailyReport.Show();
        }

       
    }
}
