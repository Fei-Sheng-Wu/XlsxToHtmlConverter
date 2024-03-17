# XlsxToHtmlConverter

[![Target Framework](https://img.shields.io/badge/.Net%20-5.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/5.0)
[![Target Framework](https://img.shields.io/badge/.Net%20Core-3.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/3.0)
[![Nuget](https://img.shields.io/badge/Nuget-v1.2.13-blue.svg?style=flat-square)](https://www.nuget.org/packages/XlsxToHtmlConverter/1.2.13)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg?style=flat-square)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.13/LICENSE.txt)

> A Xlsx to Html file converter and parser. Support cell fill, font, border, alignment and other styles. Support custom column width and row height. Support vertical and/or horizontal merged cells. Support sheet tab color and hidden sheet. Support pictures. Support progress callback. It targets both .Net 5.0 and .Net Core 3.0, and only depends on the Open Xml SDK.

## Dependencies

**DocumentFormat.OpenXml** = 2.10.1

## Main Features

- [x] Cell fill, font, border, alignment, and other styles
- [x] Custom column width and row height
- [x] Vertical and/or horizontal merged cells
- [x] Sheet tab color and hidden sheet
- [x] Pictures embedding as Base64 images
- [x] Conversion progress callback

### Original Xlsx File

![Original Xlsx File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/59e1dae0c66f3526653d82dedbe538164f19e6d2/screenshot-xlsx.png)

### Converted Html File

![Converted Html File](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/59e1dae0c66f3526653d82dedbe538164f19e6d2/screenshot-html.png)

## How to Use

Only one line to convert a xlsx file to html with the use of `Stream`.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(inputStream, outputStream);
```

Or to convert with specific `ConverterConfig` and progress callback.

```c#
XlsxToHtmlConverter.Converter.ConvertXlsx(inputStream, outputStream, config, progressCallback);
```

### Convert Local Files

A local xlsx file can be opened and read into a `MemoryStream`, and a `FileStream` can be used to write the output html into a local file.

```c#
using (MemoryStream inputStream = new MemoryStream())
{
    byte[] byteArray = File.ReadAllBytes(xlsxFileName);
    inputStream.Write(byteArray, 0, byteArray.Length);

    using FileStream outputStream = new FileStream(htmlFileName, FileMode.Create)
    {
        XlsxToHtmlConverter.Converter.ConvertXlsx(inputStream, outputStream);
    }
}
```

### Conversion Configurations

`ConverterConfig` include flexible and customizable conversion configurations.

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
}
```

### Progress Callback

A progress callback event can be set up with a `EventHandler<ConverterProgressCallbackEventArgs>`, where things like `ProgressPercent` can be used.

```c#
EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> progressCallback = null;
progressCallback += ConverterProgressCallback;
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

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.13/LICENSE.txt).
