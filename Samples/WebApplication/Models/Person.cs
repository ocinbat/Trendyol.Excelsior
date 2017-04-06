using System;
using Trendyol.Excelsior;

namespace WebApplication.Models
{
    public class Person : IExcelRow
    {
        public int RowNumber { get; set; }

        [ExcelColumn(1, "Sicil No")]
        public long Id { get; set; }

        [ExcelColumn(2, "İsim")]
        public string Name { get; set; }

        [ExcelColumn(3, "Maaş")]
        public decimal Salary { get; set; }

        [ExcelColumn(4, "İşe Giriş Tarihi", null, "dd/MM/yyyy HH:mm")]
        public DateTime EmploymentStartDate { get; set; }
    }
}