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
    /// Логика взаимодействия для ColorArticlesPage.xaml
    /// </summary>
    public partial class ColorArticlesPage : UserControl
    {
        BigFishBDEntities db;
        public List<ColorArticles> colorArticles { get; set; }
        public List<Articles> articles { get; set; }
        public ColorArticlesPage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();
            LoadArticlesData();

        }

        private void LoadArticlesData()
        {
            articles = db.Articles.ToList();
            tbTypeFK.ItemsSource = articles;
            DataContext = this;


        }
        private void combobind()
        {

            colorArticles = db.ColorArticles.ToList();
            cbSpisok.ItemsSource = colorArticles;
            cbSpisok.DisplayMemberPath = "ColorArticle";
            DataContext = this;
        }

        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            TablColorArticles.ItemsSource = colorArticles.Where(x => x.ColorArticle.Contains(tbPoisk.Text)).ToList();

        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ColorArticles cl = new ColorArticles();
                cl.ColorArticle = tbColorArt.Text;
                cl.Article = tbTypeFK.Text;
                db.ColorArticles.Add(cl);
                db.SaveChanges();
                TablColorArticles.ItemsSource = db.ColorArticles.ToList();
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
                string SDN = tbColorArt.Text;
                var SelectDN = db.ColorArticles.Where(m => m.ColorArticle == SDN).FirstOrDefault();
                db.ColorArticles.Remove(SelectDN);
                db.SaveChanges();
                TablColorArticles.ItemsSource = db.ColorArticles.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись (возможно, в главной таблице есть запись с этим артикулом)");
            }


        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablColorArticles.ItemsSource = db.ColorArticles.ToList();
            cbSpisok.ItemsSource = db.ColorArticles.ToList();
        }

        private void cbSpisok_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = (ColorArticles)cbSpisok.SelectedItem;
            TablColorArticles.ItemsSource = colorArticles.Where(x => x.ColorArticle == selected.ColorArticle).ToList();
        }

        private void tbTypeFK_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var article = item as Articles;
                return article.Article.ToLower().Contains(combo.Text.ToLower());
            };

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
