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
        public static void SaveAs(this IEnumerable<FinDataObject> list, string format)
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
                    break;
                case "csv":
                    IWriter csvWriter = new CSVWriter();
                    csvWriter.Write(path, name, list);
                    break;
            }
        }
    }
}
