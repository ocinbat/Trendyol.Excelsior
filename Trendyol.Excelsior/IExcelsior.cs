using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Trendyol.Excelsior
{
    public interface IExcelsior
    {
        IEnumerable<T> Listify<T>(string filePath, bool hasHeaderRow = false);

        IEnumerable<T> Listify<T>(IWorkbook workbook, bool hasHeaderRow = false);
    }
}