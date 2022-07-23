using FinancialData_ExcelHandler;
using FinancialData_ExcelHandler.Models;

class Program
{
    public static void Main()
    {
        string path = @"C:\Users\markell\Desktop\Телеком-Сервис ИТ\finExample.xlsx";
        Excel excel = new Excel(path, 1);

        var data = excel
            .ReadFileWithFilteringAsync(filteringNeeded: true).Result
            .ToFinDataDTO();

        foreach (var item in data)
        {
            Console.WriteLine($"Id : {item.Id}  Product : {item.Product}  Country : {item.Country}  Date : {item.Date}  Profit : {item.Profit}");
        }

        InitiateSaving(data);

        Console.ReadLine();
    }
    
    /// <summary>
    /// Запускает процесс сохранения данных в формате json /csv
    /// </summary>
    /// <param name="data"></param>
    private static void InitiateSaving(List<FinDataDTO> data, bool firstIteration = true)
    {
        if(firstIteration)
            Console.WriteLine("\nДля сохранения документа в нужном формате, введите наименование формата (в любом регистре): json / csv.");
        
        var format = Console.ReadLine();

        if (format is not null && format.ToLower() == "json" || format.ToLower() == "csv")
            data.SaveAs(format);
        else
        {
            Console.Clear();
            Console.WriteLine("Требуется ввести корректный формат сохранения данных: json или csv (в любом регистре).");
            InitiateSaving(data, false);
        }
    }
}