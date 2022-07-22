using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialData_ExcelHandler.WritingInFiles
{
    public class JSONWriter : IWriter
    {
        public void Write(string filePath, string fileName, IEnumerable<FinDataObject> list)
        {
            try
            {
                string json = JsonConvert.SerializeObject(list, Formatting.Indented);
                File.WriteAllText($"{filePath}{fileName}.json", json);
                Console.WriteLine($"Файл успешно сохранен в формате JSON по пути {filePath}{fileName}.json");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}
