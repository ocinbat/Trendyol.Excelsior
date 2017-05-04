using System;
using NPOI.SS.UserModel;
using Trendyol.Excelsior;

namespace WebApplication.Models
{
    public class Person : IExcelRow
    {
        public int RowNumber { get; set; }

        [ExcelColumn(1, "Sicil No", CellType = CellType.Numeric)]
        public long Id { get; set; }

        [ExcelColumn(2, "İsim")]
        public string Name { get; set; }

        [ExcelColumn(3, "Decimal", null, "0.00", CellType = CellType.Numeric)]
        public decimal Decimal { get; set; }

        [ExcelColumn(4, "Float", null, "0.00", CellType = CellType.Numeric)]
        public float Float { get; set; }

        [ExcelColumn(5, "Double", null, "0.00", CellType = CellType.Numeric)]
        public double Double { get; set; }

        [ExcelColumn(6, "İşe Giriş Tarihi", null, "dd/MM/yyyy HH:mm:ss")]
        public DateTime EmploymentStartDate { get; set; }
    }
}