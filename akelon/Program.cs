using akelon._data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace akelon
{
    internal class Program
    {
        static void Main(string[] args)
        {
            const string CommandGetOrderInfoByProductName = "1";
            const string CommandChangeClientContactPerson = "2";
            const string CommandGetGoldenClient = "3";
            const string CommandExit = "4";

            bool isWork = true;
            string filePath = "";

            while (true)
            {
                Console.WriteLine("Здравствуйте!\n" +
                    "Введите путь до файла Excel");

                filePath = Console.ReadLine();

                if (File.Exists(filePath) && Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    break;
                else
                    Console.WriteLine("Файл по указанному пути не найден или не является файлом Excel (.xlsx). " +
                        "Пожалуйста, введите корректный путь.");
            }

            while (isWork)
            {
                Console.WriteLine($"\n\nЧто делаем дальше?" +
                    $"\nМеню приложения следующие:" +
                    $"\nВведите {CommandGetOrderInfoByProductName} для вывода по наименованию товара информацию о клиентах" +
                    $"\nВведите {CommandChangeClientContactPerson} для изменения контактного лица Организации" +
                    $"\nВведите {CommandGetGoldenClient} для поиска золотого клиента" +
                    $"\nВведите {CommandExit} для выхода" +
                    $"\n\nВведите команду");

                string userMessage = Console.ReadLine();

                switch (userMessage)
                {
                    case CommandGetOrderInfoByProductName:
                        Console.WriteLine("Введите название продукта");
                        string productName = Console.ReadLine();

                        DataService.GetOrderInfoByProductName(filePath, productName);
                        break;

                    case CommandChangeClientContactPerson:
                        Console.WriteLine("Введите Название организации");
                        string organizationName = Console.ReadLine();

                        Console.WriteLine("Введите новое контактное лицо");
                        string newContactPerson = Console.ReadLine();

                        DataService.ChangeClientContactPerson(filePath, organizationName, newContactPerson);
                        break;

                    case CommandGetGoldenClient:
                        int year = GetValidYearFromUserInput();
                        int? month = GetValidMonthFromUserInput();

                        List<string> goldenClients = DataService.GetGoldenClients(filePath, year, month);

                        PrintGoldenClientInformation(year, month, goldenClients);

                        break;

                    case CommandExit:
                        isWork = false;
                        break;
                }
            }
        }

        private static int GetValidYearFromUserInput()
        {
            Console.WriteLine("Введите год, за который необходимо найти золотого клиента:");
            int year;

            while (!int.TryParse(Console.ReadLine(), out year) || year < 0)
            {
                Console.WriteLine("Пожалуйста, введите корректный год (целое положительное число):");
            }

            return year;
        }

        private static int? GetValidMonthFromUserInput()
        {
            Console.WriteLine("Введите месяц (необязательно, нажмите Enter для пропуска):");
            int? month = null;
            int parsedMonth;
            string monthInput = Console.ReadLine();

            if (!string.IsNullOrEmpty(monthInput))
            {
                while (!int.TryParse(monthInput, out parsedMonth) || parsedMonth < 1 || parsedMonth > 12)
                {
                    Console.WriteLine("Пожалуйста, введите корректный номер месяца (от 1 до 12):");
                    monthInput = Console.ReadLine();
                }

                month = parsedMonth;
            }

            return month;
        }

        private static void PrintGoldenClientInformation(int year, int? month, List<string> goldenClients)
        {
            if (goldenClients != null && goldenClients.Any())
            {
                Console.WriteLine($"Золотые клиенты за {year} год{(month.HasValue ? $" и {month} месяц" : string.Empty)}:");

                foreach (var client in goldenClients)
                {
                    Console.WriteLine(client);
                }
            }
            else
            {
                Console.WriteLine($"Золотые клиенты за {year} год{(month.HasValue ? $" и {month} месяц" : string.Empty)} не найдены.");
            }
        }
    }
}
