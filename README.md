# FinancialData_ExcelHandler
A console tool to iterate through the Excel file of the local PC. 

1. Reading process is implemented via a COM library Microsoft Excel Object Library.  
2. A programm Saves a read document to the FinDataModel type through the async method ReadFileWithFilteringAsync(). In the method an ability to filter a data is optional. 
3. Since a request was to take only partial data from the given Excel file, a type of FinDataDTO is created. It takes only object, that are appropriate to the request and can be transfered to the user.
4. Saving a data to the Json file is implemented via Newtonsoft.JSON NuGet-package. Saving a data to the CSV is implemented via CsvHelper NuGet-package.
5. To choose from two different formats to save a data, their implementations were split up to two different classes, both of which are implementing a parenting-imterface IWriter.
6. After saving in one of the format, a user is given a choice to save a data in the second format as well. 
