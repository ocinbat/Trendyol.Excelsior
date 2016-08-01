using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Trendyol.Excelsior
{
    public interface IExcelsior
    {
        IEnumerable<T> Listify<T>(string filePath, bool hasHeaderRow = false);

        IEnumerable<T> Listify<T>(IWorkbook workbook, bool hasHeaderRow = false);

        IEnumerable<T> Listify<T>(string filePath, IRowValidator<T> rowValidator, out IEnumerable<T> invalids, bool hasHeaderRow = false);

        IEnumerable<T> Listify<T>(IWorkbook workbook, IRowValidator<T> rowValidator, out IEnumerable<T> invalids, bool hasHeaderRow = false);

        IEnumerable<string[]> Listify(string filePath, bool hasHeaderRow = false);

        IEnumerable<string[]> Listify(IWorkbook workbook, bool hasHeaderRow = false);

        byte[] Excelify<T>(IEnumerable<T> rows, bool printHeaderRow = false);
    }
}