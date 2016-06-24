using System;
using System.Collections;

namespace Trendyol.Excelsior
{
    public class ExcelsiorException : Exception
    {
        public ExcelsiorException(string message)
            : base(message)
        {
        }
    }
}