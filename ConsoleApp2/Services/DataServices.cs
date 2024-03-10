using ClosedXML.Excel;
using ConsoleApp2.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2.Services
{
    public interface IDataService
    {
        void FindCustomersByProduct(string productName);
        void FindGoldenCustomerByProduct(int year, int month);
    }
    public  class DataServices: IDataService
    {
        private List<Order> orders;
        private List<Customer> customers;
        private List<Product> products;
        public void InitData(List<Order> orders, List<Customer> customers, List<Product> products)
        {
            this.orders = orders;
            this.customers = customers;
            this.products = products;
        }

        public void FindCustomersByProduct(string productName)
        {
            //Получить ProductCode из списка Products по наименованию товара productName
            var product = products.FirstOrDefault(p => p.Name == productName);
            if (product == null)
            {
                Console.WriteLine($"Товар с наименованием '{productName}' не найден.");
                return;
            }
            var productCode = product.ProductCode;

            //Получить из списка Order код клиента по выбранному товару
            var customerCodes = orders
                .Where(o => o.ProductCode == productCode)
                .Select(o => o.CustomerCode)
                .Distinct()
                .ToList();

            //Получить из списка Customer ContactPerson по коду клиента
            foreach (var customerCode in customerCodes)
            {
                var customer = customers.FirstOrDefault(c => c.CustomerCode == customerCode);
                if (customer != null)
                {
                    Console.WriteLine($"Клиент: {customer.OrganizationName}, Контактное лицо: {customer.ContactPerson}");
                }
            }
        }

        public void FindGoldenCustomerByProduct(int year, int month)
        {
            // Фильтрация заказов по товару и дате
            var productOrders = orders.Where(o => o.OrderDate.Year == year && o.OrderDate.Month == month).ToList();

            // Группировка заказов по коду клиента и подсчет количества заказов для каждого клиента
            var customerOrdersCount = productOrders.GroupBy(o => o.CustomerCode)
                                                    .Select(g => new { CustomerCode = g.Key, Count = g.Count() })
                                                    .OrderByDescending(x => x.Count)
                                                    .FirstOrDefault();

            if (customerOrdersCount != null)
            {
                // Поиск информации о "золотом" клиенте по его коду
                var goldenCustomer = customers.FirstOrDefault(c => c.CustomerCode == customerOrdersCount.CustomerCode);

                if (goldenCustomer != null)
                {
                    Console.WriteLine($"Золотой клиент: {goldenCustomer.OrganizationName}, Контактное лицо: {goldenCustomer.ContactPerson}, Количество заказов: {customerOrdersCount.Count}");
                }
                else
                {
                    Console.WriteLine("Информация о клиенте не найдена.");
                }
            }
            else
            {
                Console.WriteLine("Заказы на указанный товар за указанный период не найдены.");
            }
        }
    }
}
