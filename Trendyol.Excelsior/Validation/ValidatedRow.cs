using System.Collections.Generic;

namespace Trendyol.Excelsior.Validation
{
    public class ValidatedRow<T> : IValidatedRow<T>
    {
        public int RowNumber { get; set; }

        public T Item { get; set; }

        public bool IsValid { get; set; }

        public IEnumerable<IValidationError> Errors { get; set; }
    }
}