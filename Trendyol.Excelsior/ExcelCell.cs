using NPOI.SS.UserModel;

namespace Trendyol.Excelsior
{
    internal class ExcelCell
    {
        public string Value { get; set; }

        public CellType Type { get; set; }
    }
}