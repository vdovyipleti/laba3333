using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class Commodity
    {
        public int Id { get; set; }
        public string Product { get; set; }
        public int Count { get; set; }
        public double Price { get; set; }
        public double Sum { get; set; }
    }
}
