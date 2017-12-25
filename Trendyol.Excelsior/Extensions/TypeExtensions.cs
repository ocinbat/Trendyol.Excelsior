﻿using System;

namespace Trendyol.Excelsior.Extensions
{
    internal static class TypeExtensions
    {
        public static Type GetUnderlyingTypeIfPossible(this Type value)
        {
            if (value.IsGenericType && value.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                return Nullable.GetUnderlyingType(value);
            }

            return value;
        }
    }
}
