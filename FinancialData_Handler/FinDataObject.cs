using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialData_ExcelHandler
{
    public class FinDataObject
    {
        //public int Id { get; set; }
        public string Id { get; set; } // Int
        public string Segment { get; set; }
        public string Country { get; set; }
        public string Product { get; set; }
        public string DiscountBand { get; set; }

        //public float UnitsSold { get; set; }
        //public decimal ManufacturingPrice { get; set; }
        //public decimal SalePrice { get; set; }
        //public decimal GrossSales { get; set; }
        //public decimal Discounts { get; set; }
        //public decimal Sales { get; set; }
        //public decimal COGS { get; set; }
        //public decimal Profit { get; set; }
        public string UnitsSold { get; set; }
        public string ManufacturingPrice { get; set; }
        public string SalePrice { get; set; }
        public string GrossSales { get; set; }
        public string Discounts { get; set; }
        public string Sales { get; set; }
        public string COGS { get; set; }
        public string Profit { get; set; }
        public string Date { get; set; }
        //public byte MonthNumber { get; set; }
        public string MonthNumber { get; set; }
        public string MonthName { get; set; }
        //public short Year { get; set; }
        public string Year { get; set; }
    }
}
