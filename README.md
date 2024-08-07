# XlsxToHtmlConverter v1.2.17

[![C#](https://img.shields.io/badge/C%23-100%25-blue.svg?style=flat-square)](https://docs.microsoft.com/en-us/dotnet/csharp/)
[![Target Framework](https://img.shields.io/badge/.Net-%E2%89%A55.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/5.0)
[![Target Framework](https://img.shields.io/badge/.Net%20Core-%E2%89%A53.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/3.0)
[![Target Framework](https://img.shields.io/badge/.Net%20Standard-%E2%89%A52.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/platform/dotnet-standard)
[![Nuget](https://img.shields.io/badge/Nuget-%E2%89%A519.0K%20Total%20Downloads-blue.svg?style=flat-square)](https://www.nuget.org/packages/XlsxToHtmlConverter)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg?style=flat-square)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.17/LICENSE.txt)

> A Xlsx to Html file converter and parser. Support cell fill, font, border, alignment and other styles. Support custom column width and row height. Support vertical and/or horizontal merged cells. Support sheet tab color and hidden sheet. Support pictures. Support progress callback. It only depends on the Microsoft Open Xml SDK.

## Dependencies

**DocumentFormat.OpenXml** = 2.20.0

## Main Features

- [x] Cell fill, font, border, alignment, and other styles
- [x] Custom column width and row height
- [x] Vertical and/or horizontal merged cells
- [x] Sheet tab color and hidden sheet
- [x] Pictures embedding as Base64 images
- [x] Conversion progress callback

### Original Xlsx File

![Original Xlsx File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.17/screenshot-xlsx.png)

### Converted Html File

![Converted Html File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.17/screenshot-html.png)

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

A third optional parameter can be set to decide whether to use `MemoryStream` or `FileStream`. When set to `false`, it uses a `FileStream` to read the file instead of loading the entire file into a `MemoryStream` at once, which reduces the memory usage but at the cost of slowing down the conversion significantly.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(filename, outputStream, false);
```

### Conversion Configurations

`ConverterConfig` include flexible and customizable conversion configurations.

> In cases where the converter is unable to produce the correct stylings, it is suggested to set `ConvertStyle` to `false`, which will ensure the conversion of all the content with default stylings.

```c#
XlsxToHtmlConverter.ConverterConfig config = new XlsxToHtmlConverter.ConverterConfig()
{
    PageTitle = "My Title",
    PresetStyles = "body { background-color: skyblue; } table { width: 100%; }",
    ErrorMessage = "Oh, no. An error occured.",
    Encoding = System.Text.Encoding.UTF8,
    ConvertStyle = true,
    ConvertSize = true,
    ConvertPicture = true,
    ConvertSheetNameTitle = true,
    ConvertHiddenSheet = false,
    ConvertFirstSheetOnly = false,
    ConvertHtmlBodyOnly = false
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

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.17/LICENSE.txt).
