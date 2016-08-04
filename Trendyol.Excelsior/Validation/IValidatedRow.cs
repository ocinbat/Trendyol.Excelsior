using System.Collections.Generic;

namespace Trendyol.Excelsior.Validation
{
    public interface IValidatedRow<T> : IExcelRow
    {
        T Item { get; set; }

        bool IsValid { get; set; }

        IEnumerable<IValidationError> Errors { get; set; }
    }
}