using FinancialData_ExcelHandler;
class Program
{
    public static void Main()
    {
        string path = @"C:\Users\markell\Desktop\Телеком-Сервис ИТ\finExample.xlsx";
        Excel excel = new Excel(path, 1);

        var data = excel.ReadFileWithFiltering();
        foreach (var item in data)
        {
            Console.WriteLine($"id : {item.Id}\t\tProduct : {item.Product}\t\tCountry : {item.Country}\t\tDate : {item.Date}\t\tProfit : {item.Profit}");
        }
    }    
}
