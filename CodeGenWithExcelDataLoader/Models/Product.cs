using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeGenWithExcelDataLoader.Models
{
    public class Product
    {
        public string id { get; set; }
        public string name { get; set; }
        public string product_id { get; set; }
        public string valid_from { get; set; }
        public string valid_to { get; set; }
        public string description { get; set; }

    }
}
