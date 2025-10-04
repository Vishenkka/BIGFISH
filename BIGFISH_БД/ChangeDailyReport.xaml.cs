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
using System.Windows.Shapes;

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для ChangeDailyReport.xaml
    /// </summary>
    public partial class ChangeDailyReport : Window
    {
        BigFishBDEntities db;
        public List<DailyReport> dailyReport { get; set; }
        public List<Articles> articles { get; set; }
        public List<Packers> packers { get; set; }
        public List<Foundry> foundry { get; set; }
        public List<Storage> storageName { get; set; }

        public ChangeDailyReport()
        {
            InitializeComponent();
            db = new BigFishBDEntities();
            
            LoadArticlesData();
            LoadPackersData();
            LoadStorageData();
            LoadFoundryData();
            LoadColorArticles2Data();
        }

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


        private void tbPacker_TextChanged(object sender, TextChangedEventArgs e)
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

        private void tbArticle_TextChanged(object sender, TextChangedEventArgs e)
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

        private void tbArticle2_TextChanged(object sender, TextChangedEventArgs e)
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

        private void tbStorageName_TextChanged(object sender, TextChangedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.IsDropDownOpen = true;

            var collectionView = CollectionViewSource.GetDefaultView(combo.ItemsSource);
            collectionView.Filter = item =>
            {
                var storageName = item as Storage;
                return storageName.StorageName.ToLower().Contains(combo.Text.ToLower());
            };
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

        private void ChangeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Проверяем, что ID введен и является числом
                if (!int.TryParse(tbId.Text, out int sUpN))
                {
                    MessageBox.Show("Введите корректный ID");
                    return;
                }

                // Находим запись для изменения
                var SelectUptN = db.DailyReport.FirstOrDefault(w => w.Id == sUpN);

                if (SelectUptN == null)
                {
                    MessageBox.Show("Запись с указанным ID не найдена");
                    return;
                }

                // Обновляем только те поля, которые были изменены
                if (dpDatePack.SelectedDate != null)
                {
                    SelectUptN.DatePack = dpDatePack.SelectedDate.Value;
                }

                if (!string.IsNullOrEmpty(tbPacker.Text))
                {
                    SelectUptN.FIO = tbPacker.Text;
                }

                if (!string.IsNullOrEmpty(tbArticle.Text))
                {
                    SelectUptN.ArticlePack = tbArticle.Text;
                }

                if (!string.IsNullOrEmpty(tbStorageName.Text))
                {
                    SelectUptN.StorageName = tbStorageName.Text;
                }

                if (!string.IsNullOrEmpty(tbFinePacks.Text) && float.TryParse(tbFinePacks.Text, out float finePacks))
                {
                    SelectUptN.FinePacks = (float)Math.Round(finePacks, 2);
                }

                if (!string.IsNullOrEmpty(tbcolvoPacks.Text) && float.TryParse(tbcolvoPacks.Text, out float colvoPacks))
                {
                    SelectUptN.Packs = (float)Math.Round(colvoPacks, 2);
                }

                if (dpDateFoundry.SelectedDate != null)
                {
                    SelectUptN.DateFoundry = dpDateFoundry.SelectedDate.Value;
                }

                if (!string.IsNullOrEmpty(tbFoundry.Text))
                {
                    SelectUptN.FIO_Foundry = tbFoundry.Text;
                }

                if (!string.IsNullOrEmpty(tbArticle2.Text))
                {
                    SelectUptN.ArticleFoundry = tbArticle2.Text;
                }

                if (!string.IsNullOrEmpty(tbcolvoPacks2.Text) && float.TryParse(tbcolvoPacks2.Text, out float colvoPacks2))
                {
                    SelectUptN.Packs2 = (float)Math.Round(colvoPacks2, 2);
                }

                if (!string.IsNullOrEmpty(tbFinePacksFoundry.Text) && float.TryParse(tbFinePacksFoundry.Text, out float finePacksFoundry))
                {
                    SelectUptN.FinePacksFoundry = (float)Math.Round(finePacksFoundry, 2);
                }

                db.SaveChanges();
                MessageBox.Show("Изменения сохранены!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        
    }
}
