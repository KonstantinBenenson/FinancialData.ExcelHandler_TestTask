using CsvHelper;
using System.IO;
using System.Globalization;

namespace FinancialData_ExcelHandler.WritingInFiles
{
    public class CSVWriter : IWriter
    {
        public void Write(string filePath, string fileName, IEnumerable<FinDataObject> list)
        {
            using(var streamWriter = new StreamWriter(filePath + fileName + ".csv"))
            {
                using (var writer = new CsvWriter(streamWriter, CultureInfo.InvariantCulture))
                {
                    try
                    {
                        writer.WriteRecords(list);
                        Console.WriteLine($"Файл успешно сохранен в формате CSV по пути {filePath}{fileName}.csv");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);                        
                    }                    
                }
            }            
        }
    }
}
