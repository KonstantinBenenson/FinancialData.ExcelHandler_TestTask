using FinancialData_ExcelHandler.Models;
using Newtonsoft.Json;

namespace FinancialData_ExcelHandler.WritingInFiles
{
    public class JSONWriter : IWriter
    {      
        public void Write(string filePath, string fileName, List<FinDataDTO> list)
        {
            try
            {
                string json = JsonConvert.SerializeObject(list, Formatting.Indented);
                File.WriteAllText($"{filePath}{fileName}.json", json);
                Console.WriteLine($"\nФайл успешно сохранен в формате JSON по пути {filePath}{fileName}.json");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void SaveToSecondFormat(List<FinDataDTO> list)
        {
            Console.WriteLine("\nНажмите 1, если требуется сохранить файл в формате 'csv'.");
            var input = Console.ReadLine();
            if (Int32.TryParse(input, out int result))
            {
                switch (result)
                {
                    case 1:
                        Console.Clear();
                        list.SaveAs(format: "csv");
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
