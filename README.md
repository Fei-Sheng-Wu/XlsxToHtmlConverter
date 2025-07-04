# XlsxToHtmlConverter v1.2.21

[![C#](https://img.shields.io/badge/C%23-100%25-blue.svg?style=flat-square)](#)
[![Target .Net](https://img.shields.io/badge/.Net-%E2%89%A56.0-green.svg?style=flat-square)](#)
[![Target .Net Standard](https://img.shields.io/badge/.Net%20Standard-%E2%89%A52.0-green.svg?style=flat-square)](#)
[![GitHub](https://img.shields.io/github/v/release/Fei-Sheng-Wu/XlsxToHtmlConverter?label=GitHub&style=flat-square&logo=github)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/releases/latest)
[![Nuget](https://img.shields.io/nuget/v/XlsxToHtmlConverter?label=Nuget&style=flat-square&logo=nuget)](https://www.nuget.org/packages/XlsxToHtmlConverter)
[![Nuget Downloads](https://img.shields.io/nuget/dt/XlsxToHtmlConverter?label=Nuget%20Downloads&style=flat-square&logo=nuget)](https://www.nuget.org/packages/XlsxToHtmlConverter)
[![Commits Since](https://img.shields.io/github/commits-since/Fei-Sheng-Wu/XlsxToHtmlConverter/latest?label=Commits%20Since&style=flat-square)](#)
[![Lincense](https://img.shields.io/github/license/Fei-Sheng-Wu/XlsxToHtmlConverter?label=License&style=flat-square)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/LICENSE.txt)

> A fast Xlsx to Html file converter. Support cell fills, fonts, borders, alignments, and other styles. Support different cell sizes and merged cells. Support custom number formats and basic conditions. Support multiple sheets and hidden sheets. Support embedded pictures. Support progress callbacks. Only depend on the Microsoft Open Xml SDK.

## Dependencies

**DocumentFormat.OpenXml** ≥ 3.0.0

## Main Features

- [x] Cell fills, fonts, borders, alignments, and other styles
- [x] Custom column widths and row heights
- [x] Vertical and horizontal merged cells
- [x] Number formats and basic conditional formats
- [x] Sheet tab titles, colors, and hidden sheets
- [x] Picture embeddings as Base64 images
- [x] Conversion progress callback

### Original Xlsx File

![Original Xlsx File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/screenshot-xlsx.png)

### Converted Html File

![Converted Html File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/screenshot-html.png)

## How to Use

Only one line to convert a Xlsx file to Html with the use of `Stream`.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(inputStream, outputStream);
```

Or to convert with specific `ConverterConfig` and progress callback.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(inputStream, outputStream, config, progressCallback);
```

### Convert Local Files

Just use a `string` of the path to the file instead of a `Stream` to convert a local Xlsx file.

```c#
string filename = @"C:\path\to\file.xlsx";
XlsxToHtmlConverter.Converter.ConvertXlsx(filename, outputStream);
```

A third optional parameter can be set to decide whether to use a `MemoryStream` or a `FileStream`. When set to `true`, a `MemoryStream` is used to load the entire file at once instead of reading the file with a `FileStream`.

> Please note that using a `MemoryStream` will use up significantly more memory, especially for larger files.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(filename, outputStream, true);
```

### Conversion Configurations

`ConverterConfig` include flexible and customizable conversion configurations.

> In rare cases where the converter is unable to produce the correct stylings, it is suggested to set `ConvertStyles` to `false`, which will at least ensure the conversion of all the content with default stylings.

```c#
XlsxToHtmlConverter.ConverterConfig config = new XlsxToHtmlConverter.ConverterConfig()
{
    PageTitle = "My Title",
    PresetStyles = XlsxToHtmlConverter.ConverterConfig.DefaultPresetStyles + " body { background-color: skyblue; } table { width: 100%; }",
    ErrorMessage = "An unhandled error occured during the conversion: {EXCEPTION}",
    Encoding = System.Text.Encoding.UTF8,
    BufferSize = 65536,
    ConvertStyles = true,
    ConvertSizes = true,
    ConvertNumberFormats = true,
    ConvertPictures = true,
    ConvertShapes = true,
    ConvertSheetTitles = true,
    ConvertHiddenSheets = false,
    ConvertFirstSheetOnly = false,
    ConvertHtmlBodyOnly = false,
    UseHtmlStyleClasses = true,
    UseHexColors = true,
    RoundingDigits = 2
};
```

### Progress Callback

A progress callback event can be set up with `ConverterProgressCallbackEventArgs`, where things like `ProgressPercent` can be used.

```c#
EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> progressCallback = ConverterProgressCallback;
```
```c#
private static void ConverterProgressCallback(object sender, XlsxToHtmlConverter.ConverterProgressCallbackEventArgs e)
{
    string info = string.Format("{0:##0.00}% (Sheet {1} of {2} | Row {3} of {4})", e.ProgressPercent, e.CurrentSheet, e.TotalSheets, e.CurrentRow, e.TotalRows);
    string progress = new string('█', (int)(e.ProgressPercent / 2)).PadRight(50, '░');
    Console.WriteLine(info + new string(' ', 5) + progress);
}
```

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/LICENSE.txt).
