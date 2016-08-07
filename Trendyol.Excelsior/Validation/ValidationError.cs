namespace Trendyol.Excelsior.Validation
{
    public class ValidationError : IValidationError
    {
        public int ColumnOrder { get; set; }

        public string ColumnName { get; set; }

        public string Code { get; set; }

        public string DisplayMessage { get; set; }
    }
}