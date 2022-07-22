using FinancialData_ExcelHandler.Models;
using FinancialData_ExcelHandler.WritingInFiles;

namespace FinancialData_ExcelHandler
{
    public static class ExtensionMethods
    {
        /// <summary>
        /// В качестве параметра необходимо указать нужный формат: "json" или "csv"
        /// </summary>
        /// <param name="list"></param>
        /// <param name="format"></param>
        public static void SaveAs(this List<FinDataDTO> list, string format)
        {
            Console.WriteLine("Пожалуйста, введите точный адрес директории, в которую требуется сохранить файл.\nФормат адреса С:\\DirectoryName\\");
            var path = Console.ReadLine();
            Console.WriteLine("Введите имя для сохраняемого файла (без учета формата расширения).");
            var name = Console.ReadLine();

            if (path == null || name == null)
                SaveAs(list, format);

            switch (format.ToLower())
            {
                case "json":
                    IWriter jsonWriter = new JSONWriter();
                    jsonWriter.Write(path, name, list);
                    jsonWriter.SaveToSecondFormat(list);
                    break;
                case "csv":
                    IWriter csvWriter = new CSVWriter();
                    csvWriter.Write(path, name, list);
                    csvWriter.SaveToSecondFormat(list);
                    break;
            }
        }

        public static List<FinDataDTO> ToFinDataDTO(this List<FinDataModel> list)
        {
            var listDTO = new List<FinDataDTO>();
            foreach (var item in list)
            {
                listDTO.Add(new FinDataDTO
                {
                    Id = item.Id,
                    Product = item.Product,
                    Country = item.Country,
                    Date = item.Date,
                    Profit = item.Profit
                });
            }
            return listDTO;
        }
    }
}
