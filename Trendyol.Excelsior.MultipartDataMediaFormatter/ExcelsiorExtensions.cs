using System;
using System.Collections.Generic;
using System.IO;
using MultipartDataMediaFormatter.Infrastructure;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Trendyol.Excelsior.MultipartDataMediaFormatter
{
    public static class ExcelsiorExtensions
    {
        public static IEnumerable<T> Listify<T>(this IExcelsior excelsior, HttpFile httpFile, bool hasHeaderRow = false)
        {
            string fileExtension = Path.GetExtension(httpFile.FileName) ?? String.Empty;

            IWorkbook workbook;

            using (Stream stream = new MemoryStream(httpFile.Buffer))
            {
                switch (fileExtension.ToLower())
                {
                    case ".xls":
                        workbook = new HSSFWorkbook(stream);
                        break;
                    case ".xlsx":
                        workbook = new XSSFWorkbook(stream);
                        break;
                    default:
                        throw new InvalidOperationException("Excelsior can only operate on .xsl and .xlsx files.");
                }
            }

            return excelsior.Listify<T>(workbook, hasHeaderRow);
        }

        public static IEnumerable<T> Listify<T>(this IExcelsior excelsior, HttpFile httpFile, IRowValidator<T> rowValidator, out IEnumerable<T> invalids, bool hasHeaderRow = false)
        {
            string fileExtension = Path.GetExtension(httpFile.FileName) ?? String.Empty;

            IWorkbook workbook;

            using (Stream stream = new MemoryStream(httpFile.Buffer))
            {
                switch (fileExtension.ToLower())
                {
                    case ".xls":
                        workbook = new HSSFWorkbook(stream);
                        break;
                    case ".xlsx":
                        workbook = new XSSFWorkbook(stream);
                        break;
                    default:
                        throw new InvalidOperationException("Excelsior can only operate on .xsl and .xlsx files.");
                }
            }

            return excelsior.Listify<T>(workbook, rowValidator, out invalids, hasHeaderRow);
        }

        public static IEnumerable<string[]> Listify(this IExcelsior excelsior, HttpFile httpFile, bool hasHeaderRow = false)
        {
            string fileExtension = Path.GetExtension(httpFile.FileName) ?? String.Empty;

            IWorkbook workbook;

            using (Stream stream = new MemoryStream(httpFile.Buffer))
            {
                switch (fileExtension.ToLower())
                {
                    case ".xls":
                        workbook = new HSSFWorkbook(stream);
                        break;
                    case ".xlsx":
                        workbook = new XSSFWorkbook(stream);
                        break;
                    default:
                        throw new InvalidOperationException("Excelsior can only operate on .xsl and .xlsx files.");
                }
            }

            return excelsior.Listify(workbook, hasHeaderRow);
        }
    }
}
