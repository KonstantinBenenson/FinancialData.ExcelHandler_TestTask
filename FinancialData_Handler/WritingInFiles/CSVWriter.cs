using CsvHelper;
using System.Globalization;
using FinancialData_ExcelHandler.Models;

namespace FinancialData_ExcelHandler.WritingInFiles
{
    public class CSVWriter : IWriter
    {      
        public void Write(string filePath, string fileName, List<FinDataDTO> list)
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
        public void SaveToSecondFormat(List<FinDataDTO> list)
        {
            Console.WriteLine("\nНажмите 1, если требуется сохранить файл в формате 'json'.\nНажмите любую другую кнопку, чтобы выйти из программы.");
            var input = Console.ReadLine();
            if (Int32.TryParse(input, out int result))
            {
                switch (result)
                {
                    case 1:
                        Console.Clear();
                        list.SaveAs(format: "json");
                        break;
                    default:
                        break;
                };
            }
            else
                return;
        }
    }
}
