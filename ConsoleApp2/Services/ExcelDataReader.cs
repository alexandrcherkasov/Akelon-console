using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2.Model
{
    public class ExcelDataReader
    {
        private string filePath;

        public void SetExcelFilePath(string filePath)
        {
            this.filePath = filePath;
        }

        private List<Product> ReadProducts()
        {
            var products = new List<Product>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропускаем заголовок
                {
                    var product = new Product
                    {
                        ProductCode =  row.Cell(1).Value.ToString(),
                        Name = row.Cell(2).Value.ToString(),
                        Unit = row.Cell(3).Value.ToString(),
                        Price = float.Parse(row.Cell(4).Value.ToString())
                    };
                    products.Add(product);
                }
            }
            return products;
        }

        private List<Customer> ReadCustomers()
        {
            var customers = new List<Customer>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(2);
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропускаем заголовок
                {
                    customers.Add(new Customer
                    {
                        CustomerCode = row.Cell(1).Value.ToString(),
                        OrganizationName = row.Cell(2).Value.ToString(),
                        Address = row.Cell(3).Value.ToString(),
                        ContactPerson = row.Cell(4).Value.ToString(),
                    });
                }
            }
            return customers;
        }

        private List<Order> ReadOrders()
        {
            var orders = new List<Order>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(3);
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропускаем заголовок
                {
                    orders.Add(new Order
                    {
                        OrderCode = row.Cell(1).Value.ToString(),
                        ProductCode = row.Cell(2).Value.ToString(),
                        CustomerCode = row.Cell(3).Value.ToString(),
                        OrderNumber = row.Cell(4).Value.ToString(),
                        Quantity = Int32.Parse(row.Cell(5).Value.ToString()),
                        OrderDate = Convert.ToDateTime(row.Cell(6).Value)
                    });
                }
            }
            return orders;
        }

        public (List<Product> Products, List<Customer> Customers, List<Order> Orders) ReadAllData()
        {
            var products = ReadProducts();
            var customers = ReadCustomers();
            var orders = ReadOrders();
            return (products, customers, orders);
        }
    }

}
