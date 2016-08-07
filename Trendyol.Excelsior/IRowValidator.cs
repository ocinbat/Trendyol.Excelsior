using Trendyol.Excelsior.Validation;

namespace Trendyol.Excelsior
{
    public interface IRowValidator<T>
    {
        IRowValidationResult Validate<T>(T row);
    }
}