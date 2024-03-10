using ClosedXML.Excel;
using ConsoleApp2.Model;
using ConsoleApp2.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PracticTask1
{
   
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "";
            ExcelDataReader excelReader = new ExcelDataReader();
            ExcelDataChanger excelChanger = new ExcelDataChanger();
            DataServices dataServices = new DataServices();
            while (true)
            {
                Console.WriteLine("Выберите действие:");
                Console.WriteLine("1. Запрос на ввод пути до файла с данными");
                Console.WriteLine("2. По наименованию товара выводить информацию о клиентах");
                Console.WriteLine("3. Запрос на изменение контактного лица клиента");
                Console.WriteLine("4. Запрос на определение золотого клиента");
                Console.WriteLine("5. Выход");
                Console.Write("Введите номер действия: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        // Запрос на ввод пути до файла с данными
                        Console.Write("Введите путь до файла: ");
                        filePath = Console.ReadLine();
                        while (String.IsNullOrEmpty(filePath))
                        {
                            Console.WriteLine("Путь до файла не может быть пустым или null. Пожалуйста, введите путь до файла снова.");
                            Console.Write("Введите путь до файла: ");
                            filePath = Console.ReadLine();
                        }

                        try
                        {
                            excelReader.SetExcelFilePath(filePath);
                            excelChanger.SetExcelFilePath(filePath);
                            var (products, customers, orders) = excelReader.ReadAllData();
                            dataServices.InitData(orders, customers, products);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Произошла ошибка: {ex.Message}");
                        }
                        break;
                    case "2":
                        // По наименованию товара выводить информацию о клиентах
                        Console.Write("Введите наименование товара: ");
                        string productName = Console.ReadLine();
                        while (String.IsNullOrEmpty(productName))
                        {
                            Console.WriteLine("Наименование товара не может быть пустым или null. Пожалуйста, введите наименование товара снова.");
                            Console.Write("Введите наименование товара: ");
                            productName = Console.ReadLine();
                        }

                        try
                        {

                            dataServices.FindCustomersByProduct(productName);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Произошла ошибка: {ex.Message}");
                        }
                        break;
                    case "3":
                        // Запрос на изменение контактного лица клиента
                        Console.Write("Введите название организации: ");
                        string organizationName = Console.ReadLine();
                        while (String.IsNullOrEmpty(organizationName))
                        {
                            Console.WriteLine("Название организации не может быть пустым или null. Пожалуйста, введите название организации снова.");
                            Console.Write("Введите название организации: ");
                            organizationName = Console.ReadLine();
                        }

                        Console.Write("Введите ФИО нового контактного лица: ");
                        string contactPerson = Console.ReadLine();
                        while (String.IsNullOrEmpty(contactPerson))
                        {
                            Console.WriteLine("Контактное лицо не может быть пустым или null. Пожалуйста, введите ФИО нового контактного лица снова.");
                            Console.Write("Введите ФИО нового контактного лица: ");
                            contactPerson = Console.ReadLine();
                        }

                        try
                        {
                            excelChanger.UpdateContactPerson(organizationName, contactPerson);
                            var (products, customers, orders) = excelReader.ReadAllData();
                            dataServices.InitData(orders, customers, products);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Произошла ошибка: {ex.Message}");
                        }

                        break;
                    case "4":
                        // Запрос на определение золотого клиента
                        Console.Write("Введите год: ");
                        int year;
                        while (!int.TryParse(Console.ReadLine(), out year))
                        {
                            Console.WriteLine("Неверный ввод. Пожалуйста, введите год в виде числа.");
                            Console.Write("Введите год: ");
                        }

                        Console.Write("Введите месяц: ");
                        int month;
                        while (!int.TryParse(Console.ReadLine(), out month))
                        {
                            Console.WriteLine("Неверный ввод. Пожалуйста, введите месяц в виде числа.");
                            Console.Write("Введите месяц: ");
                        }
                        try
                        {
                            dataServices.FindGoldenCustomerByProduct(year, month);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Произошла ошибка: {ex.Message}");
                        }
                        break;
                    case "5":
                        // Выход из приложения
                        return;
                    default:
                        Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                        break;
                }
            }
        }
    }
}
