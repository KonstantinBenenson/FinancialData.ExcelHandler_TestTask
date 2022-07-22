using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialData_ExcelHandler
{
    public class FinDataModel
    {
        public string Id { get; set; }
        public string Segment { get; set; }
        public string Country { get; set; }
        public string Product { get; set; }
        public string DiscountBand { get; set; }
        public string UnitsSold { get; set; }
        public string ManufacturingPrice { get; set; }
        public string SalePrice { get; set; }
        public string GrossSales { get; set; }
        public string Discounts { get; set; }
        public string Sales { get; set; }
        public string COGS { get; set; }
        public string Profit { get; set; }
        public string Date { get; set; }
        public string MonthNumber { get; set; }
        public string MonthName { get; set; }
        public string Year { get; set; }
    }
}
