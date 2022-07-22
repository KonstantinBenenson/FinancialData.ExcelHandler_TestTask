using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialData_ExcelHandler.Models
{
    public class FinDataDTO
    {
        public string Id { get; set; }
        public string Product { get; set; }
        public string Country { get; set; }
        public string Date { get; set; }
        public string Profit { get; set; }        
    }
}
