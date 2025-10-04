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
    /// Логика взаимодействия для ArticlesPage.xaml
    /// </summary>
    public partial class ArticlesPage : UserControl
    {
        BigFishBDEntities db;
        public List<Articles> articles { get; set; }
        public ArticlesPage()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            combobind();

        }

        private void combobind()
        {

            articles = db.Articles.ToList();
            cbSpisok.ItemsSource = articles;
            cbSpisok.DisplayMemberPath = "Article";
            DataContext = this;
        }

        
       
        private void tbPoisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            
            TablArticles.ItemsSource = articles.Where(x => x.Article.Contains(tbPoisk.Text)).ToList();

        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Articles cl = new Articles();
                cl.Article = tbArticle.Text;
                cl.PricePackers = Convert.ToInt32(tbPricePack.Text);
                cl.PriceFoundry = Convert.ToInt32(tbPriceFoundry.Text);
                if (tbType.Text != "")
                {
                    cl.Type = Convert.ToInt32(tbType.Text);
                }
                else {  cl.Type = null; }
                    db.Articles.Add(cl);
                db.SaveChanges();
                TablArticles.ItemsSource = db.Articles.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не заполнено одно из полей или в поле 'Стоимость' введено не число.");
            }

        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string SDN = tbArticle.Text;
                var SelectDN = db.Articles.Where(m => m.Article == SDN).FirstOrDefault();
                db.Articles.Remove(SelectDN);
                db.SaveChanges();
                TablArticles.ItemsSource = db.Articles.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Не удалось удалить запись (возможно, в главной таблице есть запись с этим артикулом)");
            }


        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            db = new BigFishBDEntities();
            TablArticles.ItemsSource = db.Articles.ToList();
            cbSpisok.ItemsSource = db.Articles.ToList();
        }

        private void cbSpisok_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = (Articles)cbSpisok.SelectedItem;
            TablArticles.ItemsSource = articles.Where(x => x.Article == selected.Article).ToList();
        }

        private void Update_Click(object sender, RoutedEventArgs e)
        {
            string sUpN = tbArtNew.Text;
            var SelectUptN = db.Articles.Where(w => w.Article == sUpN).FirstOrDefault();
            SelectUptN.PricePackers = Convert.ToInt32(tbNewPricePack.Text);
            SelectUptN.PriceFoundry = Convert.ToInt32(tbNewPriceFoundry.Text);
            if (tbNewType.Text != "")
            {

                if (tbNewType.Text != "1")
                {
                    MessageBox.Show("Введен некорректный тип!");
                }
                else
                {
                    SelectUptN.Type = Convert.ToInt32(tbNewType.Text);
                }

            }
            else
            {
                SelectUptN.Type = null;
            }
            
                db.SaveChanges();
                TablArticles.ItemsSource = db.Articles.ToList();
            

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
