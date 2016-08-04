namespace Trendyol.Excelsior.Validation
{
    public interface IValidationError
    {
        int ColumnOrder { get; set; }

        string ColumnName { get; set; }

        string Code { get; set; }

        string DisplayMessage { get; set; }
    }
}