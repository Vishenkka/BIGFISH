using System;
using System.Data.SqlClient;
using System.IO;
using System.Configuration.Install;
using System.Collections;

[System.ComponentModel.RunInstaller(true)]
public class DatabaseInstaller : Installer
{
    public override void Install(IDictionary stateSaver)
    {
        base.Install(stateSaver);

        // Получаем путь установки
        string installPath = Context.Parameters["targetdir"];
        string scriptPath = Path.Combine(installPath, "script.sql");

        // Подключение к LocalDB
        string connectionString = "Server=(localdb)\\MSSQLLocalDB;Integrated Security=True;";

        try
        {
            // Читаем скрипт
            string script = File.ReadAllText(scriptPath);

            // Выполняем скрипт
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = new SqlCommand(script, connection);
                command.ExecuteNonQuery();
            }
        }
        catch (Exception ex)
        {
            // Логируем ошибку
            string logPath = Path.Combine(installPath, "install_log.txt");
            File.WriteAllText(logPath, $"Ошибка: {ex.Message}\n{ex.StackTrace}");
            throw; // Прерываем установку
        }
    }
}