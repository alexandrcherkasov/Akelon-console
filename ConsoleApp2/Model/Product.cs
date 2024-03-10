using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2.Model
{
    public class Product
    {
        public string ProductCode { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public float Price { get; set; }
    }
}
