# XlsxToHtmlConverter v1.2.19-dev

[![C#](https://img.shields.io/badge/C%23-100%25-blue.svg?style=flat-square)](https://docs.microsoft.com/en-us/dotnet/csharp/)
[![Target Framework](https://img.shields.io/badge/.Net-%E2%89%A55.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/5.0)
[![Target Framework](https://img.shields.io/badge/.Net%20Core-%E2%89%A53.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/3.0)
[![Target Framework](https://img.shields.io/badge/.Net%20Standard-%E2%89%A52.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/platform/dotnet-standard)
[![Nuget](https://img.shields.io/badge/Nuget-v1.2.18%20%28%E2%89%A519.2K%20Total%20Downloads%29-blue.svg?style=flat-square)](https://www.nuget.org/packages/XlsxToHtmlConverter/1.2.18)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg?style=flat-square)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/LICENSE.txt)

> A Xlsx to Html file converter and parser. Support cell fills, fonts, borders, alignments, and other styles. Support cell sizes and merged cells. Support custom number formats and basic conditional formats. Support multiple sheets and hidden sheets. Support embedded pictures. Support progress callbacks. Only depend on the Microsoft Open Xml SDK.

## Dependencies

**DocumentFormat.OpenXml** ≥ 2.7.1

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

A third optional parameter can be set to decide whether to use `MemoryStream` or `FileStream`. When set to `false`, it uses a `FileStream` to read the file instead of loading the entire file into a `MemoryStream` at once, effectively reducing the memory usage for larger files but at the cost of slowing down the conversion significantly.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(filename, outputStream, false);
```

### Conversion Configurations

`ConverterConfig` include flexible and customizable conversion configurations.

> In cases where the converter is unable to produce the correct stylings, it is suggested to set `ConvertStyles` to `false`, which will ensure the conversion of all the content with default stylings.

```c#
XlsxToHtmlConverter.ConverterConfig config = new XlsxToHtmlConverter.ConverterConfig()
{
    PageTitle = "My Title",
    PresetStyles = XlsxToHtmlConverter.ConverterConfig.DefaultPresetStyles + "body { background-color: skyblue; } table { width: 100%; }",
    ErrorMessage = "An unhandled error occured during the conversion: {EXCEPTION}",
    Encoding = System.Text.Encoding.UTF8,
    BufferSize = 65536,
    ConvertStyles = true,
    ConvertSizes = true,
    ConvertNumberFormats = true,
    ConvertPictures = true,
    ConvertSheetTitles = true,
    ConvertHiddenSheets = false,
    ConvertFirstSheetOnly = false,
    ConvertHtmlBodyOnly = false,
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
    string progress = new string('█', (int)(e.ProgressPercent / 2)) + new string('░', (int)((100 - e.ProgressPercent) / 2));
    Console.WriteLine(info + new string(' ', 5) + progress);
}
```

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/master/LICENSE.txt).
