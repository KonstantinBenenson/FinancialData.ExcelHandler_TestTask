using FinancialData_ExcelHandler;
class Program
{
    public static void Main()
    {
        string path = @"C:\Users\markell\Desktop\Телеком-Сервис ИТ\finExample.xlsx";
        Excel excel = new Excel(path, 1);

        var data = excel.ReadFile(default, default).Where(x => Int32.Parse(x.Profit) > 100000);
        foreach (var item in data)
        {
            Console.WriteLine($"id : {item.Id} - Product : {item.Product} - Country : {item.Country} - Date : {item.Date} - Profit : {item.Profit}");
        }
    }    
}
