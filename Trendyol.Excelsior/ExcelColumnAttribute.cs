using System;

namespace Trendyol.Excelsior
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public int Order { get; }

        public string Name { get; }

        public ExcelColumnAttribute(int order)
            : this(order, String.Empty)
        {
        }

        public ExcelColumnAttribute(int order, string name)
        {
            Order = order;
            Name = name;
        }
    }
}