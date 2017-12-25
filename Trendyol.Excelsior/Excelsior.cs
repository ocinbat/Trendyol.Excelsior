﻿using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

using Trendyol.Excelsior.Extensions;
using Trendyol.Excelsior.Validation;

namespace Trendyol.Excelsior
{
    public class Excelsior : IExcelsior
    {
        public IEnumerable<T> Listify<T>(string filePath, bool hasHeaderRow = false)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath");
            }

            string fileExtension = Path.GetExtension(filePath);

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                IWorkbook workbook;

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

                return Listify<T>(workbook, hasHeaderRow);
            }
        }

        public IEnumerable<T> Listify<T>(byte[] data, bool hasHeaderRow = false)
        {
            if (data == null || data.Length == 0)
            {
                throw new ArgumentNullException("data");
            }

            IWorkbook workbook = GetWorkbook(data);
            return Listify<T>(workbook, hasHeaderRow);
        }

        public IEnumerable<T> Listify<T>(IWorkbook workbook, bool hasHeaderRow = false)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("workbook");
            }

            ISheet sheet = workbook.GetSheetAt(0);

            if (sheet == null)
            {
                return Enumerable.Empty<T>();
            }

            int excelColumnCount = sheet.GetRow(0).PhysicalNumberOfCells;

            List<PropertyInfo> mappingTypeProperties = GetMappingTypeProperties(typeof(T));

            if (mappingTypeProperties.Count != excelColumnCount)
            {
                throw new ExcelsiorException("Column count in given file does not match columns identified in model T.");
            }

            int firstDataRow = 0;

            if (hasHeaderRow)
            {
                bool isRightColumnOrder = CheckExcelColumnOrder<T>(sheet.GetRow(0));

                if (!isRightColumnOrder)
                {
                    throw new ExcelsiorException("Columns order identified in model T does not match the order in file.");
                }

                firstDataRow = 1;
            }

            ICollection<T> itemList = new List<T>();

            for (int i = firstDataRow; i < sheet.PhysicalNumberOfRows; i++)
            {
                IRow row = sheet.GetRow(i);

                if (!IsRowEmpty(row))
                {
                    T item = GetItemFromRow<T>(row, mappingTypeProperties);

                    if (!item.Equals(default(T)))
                    {
                        itemList.Add(item);
                    }
                }
            }

            return itemList;
        }

        public IEnumerable<IValidatedRow<T>> Listify<T>(string filePath, IRowValidator<T> rowValidator, bool hasHeaderRow = false)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath");
            }

            string fileExtension = Path.GetExtension(filePath);

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                IWorkbook workbook;

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

                return Listify(workbook, rowValidator, hasHeaderRow);
            }
        }

        public IEnumerable<IValidatedRow<T>> Listify<T>(byte[] data, IRowValidator<T> rowValidator, bool hasHeaderRow = false)
        {
            if (data == null || data.Length == 0)
            {
                throw new ArgumentNullException("data");
            }

            IWorkbook workbook = GetWorkbook(data);
            return Listify<T>(workbook, rowValidator, hasHeaderRow);
        }

        public IEnumerable<IValidatedRow<T>> Listify<T>(IWorkbook workbook, IRowValidator<T> rowValidator, bool hasHeaderRow = false)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("workbook");
            }

            ISheet sheet = workbook.GetSheetAt(0);

            if (sheet == null)
            {
                return Enumerable.Empty<ValidatedRow<T>>();
            }

            int excelColumnCount = sheet.GetRow(0).PhysicalNumberOfCells;

            List<PropertyInfo> mappingTypeProperties = GetMappingTypeProperties(typeof(T));

            if (mappingTypeProperties.Count != excelColumnCount)
            {
                throw new ExcelsiorException("Column count in given file does not match columns identified in model T.");
            }

            int firstDataRow = 0;

            if (hasHeaderRow)
            {
                bool isRightColumnOrder = CheckExcelColumnOrder<T>(sheet.GetRow(0));

                if (!isRightColumnOrder)
                {
                    throw new ExcelsiorException("Columns order identified in model T does not match the order in file.");
                }

                firstDataRow = 1;
            }

            ICollection<IValidatedRow<T>> itemList = new List<IValidatedRow<T>>();

            for (int i = firstDataRow; i < sheet.PhysicalNumberOfRows; i++)
            {
                IRow row = sheet.GetRow(i);

                if (!IsRowEmpty(row))
                {
                    T item = GetItemFromRow<T>(row, mappingTypeProperties);

                    IRowValidationResult validationResult = rowValidator.Validate(item);

                    IValidatedRow<T> validatedRow = new ValidatedRow<T>();
                    validatedRow.RowNumber = firstDataRow;
                    validatedRow.Item = item;
                    validatedRow.IsValid = validationResult.IsValid;
                    validatedRow.Errors = validationResult.Errors;

                    itemList.Add(validatedRow);
                }
            }

            return itemList;
        }

        public IEnumerable<string[]> Arrayify(string filePath, bool hasHeaderRow = false)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath");
            }

            string fileExtension = Path.GetExtension(filePath);

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                IWorkbook workbook;

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

                return Arrayify(workbook, hasHeaderRow);
            }
        }

        public IEnumerable<string[]> Arrayify(byte[] data, bool hasHeaderRow = false)
        {
            if (data == null || data.Length == 0)
            {
                throw new ArgumentNullException("data");
            }

            IWorkbook workbook = GetWorkbook(data);
            return Arrayify(workbook, hasHeaderRow);
        }

        public IEnumerable<string[]> Arrayify(IWorkbook workbook, bool hasHeaderRow = false)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("workbook");
            }

            ISheet sheet = workbook.GetSheetAt(0);

            if (sheet == null)
            {
                return Enumerable.Empty<string[]>();
            }

            int firstDataRow = 0;

            if (hasHeaderRow)
            {
                firstDataRow = 1;
            }

            ICollection<string[]> itemList = new List<string[]>();
            int maxColumnNumber = GetMaxColumnNumberOfSheet(sheet);
            int maxRowNumber = GetMaxRowNumberOfSheet(sheet);

            for (int i = firstDataRow; i < maxRowNumber + 1; i++)
            {
                IRow row = sheet.GetRow(i);
                string[] itemRow = GetItemFromRow(row, maxColumnNumber);
                itemList.Add(itemRow);
            }

            return itemList;
        }

        public byte[] Excelify<T>(IEnumerable<T> rows, bool printHeaderRow = false)
        {
            if (!printHeaderRow && (rows == null || !rows.Any()))
            {
                throw new ArgumentException("There must be at least one row to generate excel file.", "rows");
            }

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");
            int rowIndex = 0;

            List<PropertyInfo> mappingTypeProperties = GetMappingTypeProperties(typeof(T));
            mappingTypeProperties = mappingTypeProperties.OrderBy(p => p.GetCustomAttribute<ExcelColumnAttribute>().Order).ToList();

            if (printHeaderRow)
            {
                IRow headerRow = sheet.CreateRow(rowIndex++);

                List<string> headerColumns = mappingTypeProperties
                    .Select(p => p.GetCustomAttribute<ExcelColumnAttribute>().Name)
                    .ToList();

                for (int i = 0; i < headerColumns.Count; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(headerColumns[i]);
                }

                SetHeaderRowStyles(workbook, headerRow);
            }

            List<T> items = rows.ToList();

            int? cellCount = null;

            for (int i = 0; i < items.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(rowIndex++);

                List<ExcelCell> rowCells = GetCellArrayForItem(items[i], mappingTypeProperties);

                if (!cellCount.HasValue)
                {
                    cellCount = rowCells.Count;
                }

                for (int j = 0; j < rowCells.Count; j++)
                {
                    ICell cell = dataRow.CreateCell(j, rowCells[j].Type);

                    switch (rowCells[j].Type)
                    {
                        case CellType.Numeric:
                            cell.SetCellValue(Convert.ToDouble(rowCells[j].Value));

                            if (!String.IsNullOrEmpty(rowCells[j].Format))
                            {
                                ICellStyle style = workbook.CreateCellStyle();
                                style.DataFormat = HSSFDataFormat.GetBuiltinFormat(rowCells[j].Format);
                                cell.CellStyle = style;
                            }
                            break;
                        case CellType.String:
                            cell.SetCellValue(rowCells[j].Value == null ? String.Empty : $"{rowCells[j].Value}");
                            break;
                        default:
                            throw new ArgumentOutOfRangeException("CellType", $"Unsupported cell type: {rowCells[j].Type}.");
                    }
                }
            }

            if (cellCount.HasValue)
            {
                for (int i = 0; i < cellCount.Value; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }

            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Write(stream);

                return stream.ToArray();
            }
        }

        private void SetHeaderRowStyles(IWorkbook workbook, IRow headerRow)
        {
            IFont font = CreateHeaderFont(workbook);
            ICellStyle style = workbook.CreateCellStyle();
            style.SetFont(font);
            style.FillForegroundColor = HSSFColor.Grey50Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;

            headerRow.Cells.ForEach(c => c.CellStyle = style);
        }

        private IFont CreateHeaderFont(IWorkbook workbook)
        {
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 13;
            font.FontName = "Calibri";
            font.Color = HSSFColor.White.Index;
            font.Boldweight = (short)FontBoldWeight.Bold;
            return font;
        }

        private List<PropertyInfo> GetMappingTypeProperties(Type type)
        {
            return type.GetProperties().Where(p => p.GetCustomAttribute<ExcelColumnAttribute>() != null).ToList();
        }

        private bool CheckExcelColumnOrder<T>(IRow headerRow)
        {
            List<PropertyInfo> properties = GetMappingTypeProperties(typeof(T));

            if (headerRow != null)
            {
                for (int i = 0; i < headerRow.Cells.Count; i++)
                {
                    string cellValue = headerRow.Cells[i].StringCellValue;

                    PropertyInfo pi = properties.FirstOrDefault(p => p.GetCustomAttribute<ExcelColumnAttribute>().Name == cellValue);

                    if (pi == null)
                    {
                        return false;
                    }

                    if (pi.GetCustomAttribute<ExcelColumnAttribute>().Order != i + 1)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private bool IsRowEmpty(IRow row)
        {
            if (row == null)
            {
                return true;
            }

            using (var enumerator = row.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ICell cell = enumerator.Current;

                    if (cell != null && cell.CellType != CellType.Blank)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private int GetMaxRowNumberOfSheet(ISheet sheet)
        {
            int result = 0;

            var enumerator = sheet.GetRowEnumerator();

            while (enumerator.MoveNext())
            {
                IRow row = (IRow)enumerator.Current;

                if (row != null && row.RowNum > result)
                {
                    result = row.RowNum;
                }
            }

            return result;
        }

        private T GetItemFromRow<T>(IRow row, List<PropertyInfo> mappingTypeProperties)
        {
            var item = Activator.CreateInstance<T>();

            foreach (PropertyInfo pi in mappingTypeProperties)
            {
                var attr = pi.GetCustomAttribute<ExcelColumnAttribute>();

                if (attr != null)
                {
                    int columnOrder = attr.Order - 1;

                    string columnValue = row.GetCell(columnOrder, MissingCellPolicy.RETURN_BLANK_AS_NULL) == null ? string.Empty : row.GetCell(columnOrder, MissingCellPolicy.RETURN_BLANK_AS_NULL).ToString();

                    if (!string.IsNullOrEmpty(columnValue))
                    {
                        columnValue = columnValue.Trim();

                        if (pi.PropertyType.GetUnderlyingTypeIfPossible() == typeof(int))
                        {
                            int val;

                            if (int.TryParse(columnValue, out val))
                            {
                                pi.SetValue(item, val);
                            }
                        }
                        else if (pi.PropertyType.GetUnderlyingTypeIfPossible() == typeof(decimal))
                        {
                            decimal val;

                            if (decimal.TryParse(columnValue, out val))
                            {
                                pi.SetValue(item, Math.Round(val, 2));
                            }
                        }
                        else if (pi.PropertyType.GetUnderlyingTypeIfPossible() == typeof(float))
                        {
                            float val;

                            if (float.TryParse(columnValue, out val))
                            {
                                pi.SetValue(item, Math.Round(val, 2));
                            }
                        }
                        else if (pi.PropertyType.GetUnderlyingTypeIfPossible() == typeof(long))
                        {
                            long val;

                            if (long.TryParse(columnValue, out val))
                            {
                                pi.SetValue(item, val);
                            }
                        }
                        else if (pi.PropertyType.GetUnderlyingTypeIfPossible() == typeof(DateTime))
                        {
                            DateTime val;

                            if (String.IsNullOrEmpty(attr.Format))
                            {
                                if (DateTime.TryParse(columnValue, out val))
                                {
                                    pi.SetValue(item, val);
                                }
                            }
                            else
                            {
                                if (DateTime.TryParseExact(columnValue, attr.Format, CultureInfo.InvariantCulture, DateTimeStyles.None, out val))
                                {
                                    pi.SetValue(item, val);
                                }
                            }
                        }
                        else if (pi.PropertyType.GetUnderlyingTypeIfPossible() == typeof(string))
                        {
                            pi.SetValue(item, columnValue);
                        }
                    }
                    else
                    {
                        pi.SetValue(item, attr.DefaultValue);
                    }
                }
            }

            return item;
        }

        private string[] GetItemFromRow(IRow row, int columnCountOfSheet)
        {
            string[] itemRow = new string[columnCountOfSheet + 1];
            bool isRowEmpty = IsRowEmpty(row);

            if (isRowEmpty)
            {
                return itemRow;
            }

            row.Cells.ForEach(c => c.SetCellType(CellType.String));

            foreach (ICell cell in row.Cells)
            {
                if (cell != null)
                {
                    itemRow[cell.ColumnIndex] = cell.StringCellValue;
                }
            }

            return itemRow;
        }

        private int GetMaxColumnNumberOfSheet(ISheet sheet)
        {
            int result = 0;

            var enumerator = sheet.GetRowEnumerator();

            while (enumerator.MoveNext())
            {
                IRow row = (IRow)enumerator.Current;

                if (!IsRowEmpty(row))
                {
                    int rowColumCount = row.Cells.Max(c => c.ColumnIndex);

                    if (rowColumCount > result)
                    {
                        result = rowColumCount;
                    }
                }
            }

            return result;
        }

        private List<ExcelCell> GetCellArrayForItem<T>(T item, List<PropertyInfo> mappingTypeProperties)
        {
            List<ExcelCell> cells = new List<ExcelCell>();

            foreach (PropertyInfo pi in mappingTypeProperties)
            {
                ExcelCell cell = new ExcelCell();

                ExcelColumnAttribute attr = pi.GetCustomAttribute<ExcelColumnAttribute>();

                if (attr.CellType == CellType.Unknown)
                {
                    attr.CellType = CellType.String;
                }

                cell.Format = attr.Format;
                cell.Type = attr.CellType;

                object itemValue = pi.GetValue(item);

                if (itemValue != null)
                {
                    if (pi.PropertyType == typeof(DateTime))
                    {
                        if (String.IsNullOrEmpty(attr.Format))
                        {
                            cell.Value = ((DateTime)itemValue).ToString(CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            cell.Value = ((DateTime)itemValue).ToString(attr.Format, CultureInfo.InvariantCulture);
                        }
                    }
                    else
                    {
                        cell.Value = itemValue;
                    }
                }

                cells.Add(cell);
            }

            return cells;
        }

        private IWorkbook GetWorkbook(byte[] data)
        {
            IWorkbook workbook;

            if (TryCreateHSSFWorkbook(data, out workbook) || TryCreateXSSFWorkbook(data, out workbook))
            {
                return workbook;
            }

            throw new InvalidOperationException("Excelsior can only operate on .xsl and .xlsx files.");
        }

        private bool TryCreateHSSFWorkbook(byte[] data, out IWorkbook workbook)
        {
            MemoryStream stream = new MemoryStream(data);

            try
            {
                workbook = new HSSFWorkbook(stream);
                return true;
            }
            catch (Exception)
            {
                workbook = null;
            }
            finally
            {
                stream.Close();
                stream.Dispose();
            }

            return false;
        }

        private bool TryCreateXSSFWorkbook(byte[] data, out IWorkbook workbook)
        {
            MemoryStream stream = new MemoryStream(data);

            try
            {
                workbook = new XSSFWorkbook(stream);
                return true;
            }
            catch (Exception)
            {
                workbook = null;
            }
            finally
            {
                stream.Close();
                stream.Dispose();
            }

            return false;
        }
    }
}