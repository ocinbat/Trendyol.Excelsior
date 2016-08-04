using System.Collections.Generic;

namespace Trendyol.Excelsior.Validation
{
    public interface IRowValidationResult
    {
        bool IsValid { get; set; }

        IEnumerable<IValidationError> Errors { get; set; }
    }
}