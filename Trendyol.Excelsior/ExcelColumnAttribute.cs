﻿using System;

namespace Trendyol.Excelsior
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public int Order { get; }

        public string Name { get; }

        public object DefaultValue { get; }

        public string Format { get; }

        public ExcelColumnAttribute(int order)
            : this(order, String.Empty)
        {
        }

        public ExcelColumnAttribute(int order, string name)
            : this(order, name, null)
        {
        }

        public ExcelColumnAttribute(int order, string name, object defaultValue)
            : this(order, name, defaultValue, String.Empty)
        {
        }

        public ExcelColumnAttribute(int order, string name, object defaultValue, string format)
        {
            Order = order;
            Name = name;
            DefaultValue = defaultValue;
            Format = format;
        }
    }
}