# XlsxToHtmlConverter v2.0.0

[![Language](https://img.shields.io/badge/Language-C%23-lightgray.svg?style=flat-square)](#)
[![.NET](https://img.shields.io/badge/.NET-%E2%89%A56.0-orange.svg?style=flat-square)](#)
[![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%E2%89%A52.0-orange.svg?style=flat-square)](#)
[![NuGet](https://img.shields.io/nuget/v/XlsxToHtmlConverter?label=NuGet&style=flat-square&logo=nuget)](https://www.nuget.org/packages/XlsxToHtmlConverter)
[![Downloads](https://img.shields.io/nuget/dt/XlsxToHtmlConverter?label=Downloads&style=flat-square&logo=nuget)](https://www.nuget.org/packages/XlsxToHtmlConverter)
[![Commits Since](https://img.shields.io/github/commits-since/Fei-Sheng-Wu/XlsxToHtmlConverter/latest?label=Commits%20Since&style=flat-square)](#)
[![License](https://img.shields.io/github/license/Fei-Sheng-Wu/XlsxToHtmlConverter?label=License&style=flat-square)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/LICENSE.txt)

> A fast, versatile, and powerful XLSX to HTML converter. Support an extensive scope of cell stylings and additional elements. Empower the efficient transformation of spreadsheets into well-structured web documents. Provide the ability to easily customize every aspect of the conversion process with progress callbacks. Only depend on the Open XML SDK.

## Dependencies

- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) ≥ 3.0.0, < 4.0.0

## Features

- [x] Cell structures, sizes, fonts, fills, borders, alignments, and visibilities
- [x] Content presentation with number formats and basic conditional formats
- [x] Elements of pictures and shapes with responsive positioning
- [x] HTML construction with configurable details and modernized organization

### Original XLSX File

![Original XLSX File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/sample-xlsx.png)

### Converted HTML File

![Converted HTML File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/sample-html.png)

## Versioning

For versions ≥ v2.0.0, the versioning of XlsxToHtmlConverter conforms to the following scheme:

| Generation | | Major | | Minor |
| :--- | :---: | :--- | :---: | :--- |
| **2** | . | **0** | . | **0** |
| _(backward-incompatible)_ | | _(backward-incompatible)_ | | _(backward-compatible)_ |
| Significant codebase refactors. | | Severe bug fixes and core improvements. | | Mild changes. |

## How to Use

Only one line to convert a local XLSX file to HTML:

```c#
XlsxToHtmlConverter.Converter.Convert(@"C:\path\to\input.xlsx", @"C:\path\to\output.html");
```

Similarly, the use of `Stream` is supported:

```c#
Stream input = ...;
Stream output = ...;
XlsxToHtmlConverter.Converter.Convert(input, output);
```

Alternatively, the input may also be a `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument` instance:

```c#
using DocumentFormat.OpenXml.Packaging;
```
```c#
SpreadsheetDocument input = ...;
Stream output = ...;
XlsxToHtmlConverter.Converter.Convert(input, output);
```

### Conversion Configuration

A third optional parameter can be set to configure the conversion process:

```c#
XlsxToHtmlConverter.ConverterConfiguration configuration = new()
{
    BufferSize = 65536,
    Encoding = Encoding.UTF8,
    NewlineCharacter = "\n",
    RoundingDigits = 2,
    CurrentCulture = CultureInfo.CurrentCulture,
    HtmlTitle = null,
    HtmlPresetStylesheet = ...,
    XlsxSheetSelector = null,
    ConvertSheetTitles = true,
    ConvertSizes = true,
    ConvertVisibilities = true,
    ConvertStyles = true,
    ConvertNumberFormats = true,
    ConvertPictures = true,
    ConvertShapes = true,
    UseHtmlFragment = false,
    UseHtmlClasses = true,
    UseHtmlHexColors = true,
    UseHtmlProportionalWidths = true,
    UseHtmlDataElements = true,
    ...
};
XlsxToHtmlConverter.Converter.Convert(..., ..., configuration);
```

### Progress Callback

A fourth optional parameter can be set to add a progress callback event handler:

```c#
XlsxToHtmlConverter.Converter.Convert(..., ..., ..., HandleProgressChanged);
```
```c#
private void HandleProgressChanged(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument? sender, XlsxToHtmlConverter.ConverterProgressChangedEventArgs e)
{
    string summary = $"Sheet {e.CurrentSheet} of {e.SheetCount} | Row {e.CurrentRow} of {e.RowCount}";
    string progress = new string('█', (int)Math.Round(e.ProgressPercentage / 100.0 * 50)).PadRight(50, '░');
    Console.Write($"{e.ProgressPercentage:F2}% ({summary}) {progress}");
}
```

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/LICENSE.txt).
