using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
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
    /// Логика взаимодействия для Authorization.xaml
    /// </summary>
    public partial class Authorization : Window
    {
        BigFishBDEntities db;
        public Authorization()
        {
            InitializeComponent();
           
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string login = tbL.Text.Trim();
                string password = tbP.Text;

                using (var db = new BigFishBDEntities())
                {
                    var user = db.User1.FirstOrDefault(u => u.Login == login && u.Password == password);

                    if (user != null)
                    {
                        if (user.RoleNumber == 0) //окно бухгалтера
                        {
                            new MainWindow().Show();
                            this.Close();
                        }

                        else if (user.RoleNumber == 1) //окно админа
                        {
                            new AdminMainWindow().Show();
                            this.Close();
                        }

                    }
                    else
                    {
                        
                            MessageBox.Show($"Неверный логин или пароль!");
                        
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
    }
}
