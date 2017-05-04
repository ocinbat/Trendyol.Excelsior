# Trendyol.Excelsior

[![Build status](https://ci.appveyor.com/api/projects/status/s50h6ad5mip73vyw?svg=true)](https://ci.appveyor.com/project/ocinbat/trendyol-excelsior)

A wrapper library to manipulate MS Excel files easily.

## Getting Started

### Prerequisites

Trendyol.Excelsior only supports .NET 4.6.1 at this moment.

### Installing

In order to use Excelsior on your project you need to add below nuget package to your project.

```
Install-Package Trendyol.Excelsior
```

Initialize an Excelsior instance.

```csharp
IExcelsior excelsior = new Excelsior();
```

Use appropirate methods for your needs.

```csharp
public interface IExcelsior
{
    // Get a generic IEnumerable from a file.
    IEnumerable<T> Listify<T>(string filePath, bool hasHeaderRow = false);
    
    // Get a generic IEnumerable from a byte array.
    IEnumerable<T> Listify<T>(byte[] data, bool hasHeaderRow = false);

    // Get a generic IEnumerable from an NPOI Workbook.
    IEnumerable<T> Listify<T>(IWorkbook workbook, bool hasHeaderRow = false);

    // Excelsior also supports custom validation for rows.
    // All you need to do is to implement an IRowValidator<T> for your class and give an instance to Excelsior when you call desired method.
    // All rows are wrapper in IValidatedRow interface with extra properties indicating if the row is valid or not.
    IEnumerable<IValidatedRow<T>> Listify<T>(string filePath, IRowValidator<T> rowValidator, bool hasHeaderRow = false);

    IEnumerable<IValidatedRow<T>> Listify<T>(byte[] data, IRowValidator<T> rowValidator, bool hasHeaderRow = false);

    IEnumerable<IValidatedRow<T>> Listify<T>(IWorkbook workbook, IRowValidator<T> rowValidator, bool hasHeaderRow = false);

    // You can also get a string array from your file.
    IEnumerable<string[]> Arrayify(string filePath, bool hasHeaderRow = false);

    IEnumerable<string[]> Arrayify(byte[] data, bool hasHeaderRow = false);

    IEnumerable<string[]> Arrayify(IWorkbook workbook, bool hasHeaderRow = false);

    // Get a byte array. You can eitjer save this to a file or return from a controller.
    byte[] Excelify<T>(IEnumerable<T> rows, bool printHeaderRow = false);
}
```

For built-in cell formats:
[NPOI](https://github.com/tonyqus/npoi/blob/02f080d3ee37e4f04a999be32604b1cb6bf3e649/main/SS/UserModel/BuiltinFormats.cs)

## Built With

* [NPOI](https://github.com/tonyqus/npoi) - a .NET library that can read/write Office formats without Microsoft Office installed. No COM+, no interop.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
