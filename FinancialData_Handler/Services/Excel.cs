using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using _Excel = Microsoft.Office.Interop.Excel;

namespace FinancialData_ExcelHandler
{
    public class Excel
    {
        private string _path = string.Empty;
        Application excel = new Application();
        Workbook workbook;
        Worksheet worksheet;

        public Excel(string path, int sheet)
        {
            _path = path;
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheet];
        }

        /// <summary>
        /// Считывает содержание Excel файла из указанной директории.
        /// В качестве параметра принимает первую строку, с которой начнется считывание документа (по умолчанию значение = 2),
        /// а также булиевую переменную, определяющемую необходимость фильтрации по столбцу Profit > 100000 (по умолчанию False)
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public async Task<List<FinDataModel>> ReadFileWithFilteringAsync(bool filteringNeeded = false, int firstRow = 2)
        {
            var finDataObjects = new List<FinDataModel>();
            int rows = GetRowsCount();

            try
            {
                await Task.Run(() =>
                {
                    for (int i = firstRow; i <= rows; i++) // Rows
                    {
                        if (filteringNeeded)
                        {
                            Int32.TryParse(Convert.ToString(worksheet.Cells[i, 13].Value2), out int profit);
                            if (profit.GetType() == typeof(int) && !(profit >= 100000))
                                continue;
                        }
                        {
                            try
                            {
                                finDataObjects.Add(new FinDataModel()
                                {
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
                                });
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                Console.ReadLine();
                            }
                        }
                    }
                });                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                QuitAndRelease();               
            }

            return finDataObjects;
        }

        /// <summary>
        /// Получает кол-во строк в Excel-файле
        /// </summary>
        /// <returns></returns>
        private int GetRowsCount()
        {
            _Excel.Range range = worksheet.UsedRange;
            return range.Rows.Count;
        }

        /// <summary>
        /// Закрывает рабочий файл Excel, а также заканчивает выполнение подключения к Excel
        /// </summary>
        private void QuitAndRelease()
        {
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }

    }
}
