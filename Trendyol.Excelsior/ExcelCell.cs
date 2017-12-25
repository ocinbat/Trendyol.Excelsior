using NPOI.SS.UserModel;

namespace Trendyol.Excelsior
{
    internal class ExcelCell
    {
        public object Value { get; set; }

        public string Format { get; set; }

        public CellType Type { get; set; }
    }
}