namespace Trendyol.Excelsior
{
    public interface IRowValidator<T>
    {
        bool IsValid(T row);
    }
}