﻿using System.Collections.Generic;
using NPOI.SS.UserModel;
using Trendyol.Excelsior.Validation;

namespace Trendyol.Excelsior
{
    public interface IExcelsior
    {
        IEnumerable<T> Listify<T>(string filePath, bool hasHeaderRow = false);

        IEnumerable<T> Listify<T>(byte[] data, bool hasHeaderRow = false);

        IEnumerable<T> Listify<T>(IWorkbook workbook, bool hasHeaderRow = false);

        IEnumerable<IValidatedRow<T>> Listify<T>(string filePath, IRowValidator<T> rowValidator, bool hasHeaderRow = false);

        IEnumerable<IValidatedRow<T>> Listify<T>(byte[] data, IRowValidator<T> rowValidator, bool hasHeaderRow = false);

        IEnumerable<IValidatedRow<T>> Listify<T>(IWorkbook workbook, IRowValidator<T> rowValidator, bool hasHeaderRow = false);

        IEnumerable<string[]> Arrayify(string filePath, bool hasHeaderRow = false);

        IEnumerable<string[]> Arrayify(byte[] data, bool hasHeaderRow = false);

        IEnumerable<string[]> Arrayify(IWorkbook workbook, bool hasHeaderRow = false);

        byte[] Excelify<T>(IEnumerable<T> rows, bool printHeaderRow = false);
    }
}