﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Trendyol.Excelsior.Tests
{
    [TestClass]
    public class ExcelsiorTests
    {
        [TestMethod]
        public void test_for_nullable_property_parse()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            string uriPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase) + "\\spreadsheets\\GeneticTypeWorkbook.xlsx";
            string filePath = new Uri(uriPath).LocalPath;

            var excelsior = new Excelsior();

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            List<StockImportExcelDto> result = excelsior.Listify<StockImportExcelDto>(filePath, true).ToList();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            Assert.AreEqual(result.Count, 4);
            Assert.IsTrue(result.Any(x => x.Quantity.HasValue));

        }

        public class StockImportExcelDto
        {
            [ExcelColumn(1, "Barcode")]
            public string Barcode { get; set; }

            [ExcelColumn(2, "Quantity")]
            public int? Quantity { get; set; }
        }
    }
}