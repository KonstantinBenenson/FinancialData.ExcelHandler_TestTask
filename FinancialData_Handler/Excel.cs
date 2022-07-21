using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using _Excel = Microsoft.Office.Interop.Excel;

namespace FinancialData_ExcelHandler
{
    public class Excel
    {
        private string _path = string.Empty;
        _Application excel = new _Excel.Application();
        Workbook workbook;
        Worksheet worksheet;

        public Excel(string path, int sheet)
        {
            _path = path;
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheet];
        }

        /// <summary>
        /// A method receives a starting cell pointers (axises), 
        /// also a user can decide whether he wants to filter a data by the profits > 100000. By default a value of this parameter equals True 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public List<FinDataObject> ReadFile(bool filteringNeeded = true, int firstRow = 2)
        {
            var finDataObjects = new List<FinDataObject>();
            int rows = GetRowsCount();
            for (int i = firstRow; i <= rows; i++) // Rows
            {
                if(filteringNeeded)
                {
                    Int32.TryParse(Convert.ToString(worksheet.Cells[i, 13].Value2), out int profit);
                    if (profit.GetType() == typeof(int) && !(profit >= 100000))
                        continue;
                }                               
                {
                    try
                    {
                        var FDO = new FinDataObject()
                        {
                            //Id = Int32.Parse(worksheet.Cells[i, j].Value2),
                            //Segment = worksheet.Cells[i, j + 1].Value2,
                            //Country = worksheet.Cells[i, j + 2].Value2,
                            //Product = worksheet.Cells[i, j + 3].Value2,
                            //DiscountBand = worksheet.Cells[i, j + 4].Value2,
                            //UnitsSold = float.Parse(worksheet.Cells[i, j + 5].Value2),
                            //ManufacturingPrice = Convert.ToDecimal(worksheet.Cells[i, j + 6].Value2),
                            //SalePrice = Convert.ToDecimal(worksheet.Cells[i, j + 7].Value2),
                            //GrossSales = Convert.ToDecimal(worksheet.Cells[i, j + 8].Value2),
                            //Discounts = Convert.ToDecimal(worksheet.Cells[i, j + 9].Value2),
                            //Sales = Convert.ToDecimal(worksheet.Cells[i, j + 10].Value2),
                            //COGS = Convert.ToDecimal(worksheet.Cells[i, j + 11].Value2),
                            //Profit = Convert.ToDecimal(worksheet.Cells[i, j + 12].Value2),
                            //Date = worksheet.Cells[i, j + 13].Value2,
                            //MonthNumber = byte.Parse(worksheet.Cells[i, j + 14].Value2),
                            //MonthName = worksheet.Cells[i, j + 15].Value2,
                            //Year = short.Parse(worksheet.Cells[i, j + 16].Value2)
                            Id = Convert.ToString(worksheet.Cells[i, 1].Value2),
                            Segment = Convert.ToString(worksheet.Cells[i, 2].Value2),
                            Country = Convert.ToString(worksheet.Cells[i, 3].Value2),
                            Product = Convert.ToString(worksheet.Cells[i, 4].Value2),
                            DiscountBand = Convert.ToString(worksheet.Cells[i, 5].Value2),
                            UnitsSold = Convert.ToString(worksheet.Cells[i, 6].Value2),
                            ManufacturingPrice = Convert.ToString(worksheet.Cells[i, 7].Value2),
                            SalePrice = Convert.ToString(worksheet.Cells[i, 8].Value2),
                            GrossSales = Convert.ToString(worksheet.Cells[i, 9].Value2),
                            Discounts = Convert.ToString(worksheet.Cells[i, 10].Value2),
                            Sales = Convert.ToString(worksheet.Cells[i, 11].Value2),
                            COGS = Convert.ToString(worksheet.Cells[i, 12].Value2),
                            Profit = Convert.ToString(worksheet.Cells[i, 13].Value2),
                            Date = Convert.ToString(worksheet.Cells[i, 14].Value2),
                            MonthNumber = Convert.ToString(worksheet.Cells[i, 15].Value2),
                            MonthName = Convert.ToString(worksheet.Cells[i, 16].Value2),
                            Year = Convert.ToString(worksheet.Cells[i, 17].Value2)
                        };
                        finDataObjects.Add(FDO);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.ReadLine();
                    }
                }
            }
            return finDataObjects;
        }

        //private static string ClearValue (this string str)
        //{
        //    if (string.IsNullOrEmpty(str))
        //        return str;

        //}

        private int GetRowsCount()
        {
            _Excel.Range range = worksheet.UsedRange;
            return range.Rows.Count;
        }

        public void QuitAndRelease()
        {
            workbook.Close();
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }

    }
}
