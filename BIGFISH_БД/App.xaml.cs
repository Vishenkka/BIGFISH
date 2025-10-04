using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace BIGFISH_БД
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            try
            {
                Database.SetInitializer(new CreateDatabaseIfNotExists<BigFishBDEntities>());
                using (var db = new BigFishBDEntities())
                {
                    if (!db.Database.Exists()) //удалить? уже ж не надо
                    {
                        db.Database.Create();
                        MessageBox.Show("База данных успешно создана", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }

                    
                }
            }
            catch (System.Data.SqlClient.SqlException sqlEx)
            {
                MessageBox.Show($"Ошибка SQL: {sqlEx.Message}\n\nПроверьте:\n" +
                               "1. Установлен ли LocalDB\n" +
                               "2. Запущена ли служба LocalDB\n" +
                               "3. Правильность строки подключения",
                               "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Неизвестная ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
        }
    }
}
