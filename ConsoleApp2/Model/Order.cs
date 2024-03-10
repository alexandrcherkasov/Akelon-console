using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2.Model
{
    public class Order
    {
        public string OrderCode { get; set; }
        public string ProductCode { get; set; }
        public string CustomerCode { get; set; }
        public string OrderNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }
}
