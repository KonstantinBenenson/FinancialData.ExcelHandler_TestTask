using FinancialData_ExcelHandler.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialData_ExcelHandler.WritingInFiles
{
    public interface IWriter
    {
        void Write(string filePath, string fileName, List<FinDataDTO> list);
        void SaveToSecondFormat(List<FinDataDTO> list);
    }
}
