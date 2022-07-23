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
        /// <summary>
        /// Записывает список объектов FinDataDTO в формат json / csv на выбор
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <param name="list"></param>
        void Write(string filePath, string fileName, List<FinDataDTO> list);

        /// <summary>
        /// Предлагает пользователю сохранить список объектов во втором формате (в зависимости от того, какой формат уже был сохранен)
        /// </summary>
        /// <param name="list"></param>
        void SaveToSecondFormat(List<FinDataDTO> list);
    }
}
