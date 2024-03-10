# XlsxToHtmlConverter

[![Target Framework](https://img.shields.io/badge/.Net%20-6.0-green.svg?style=flat-square)](https://dotnet.microsoft.com/en-us/download/dotnet/6.0)
[![Nuget](https://img.shields.io/badge/Nuget-v1.2.7-blue.svg?style=flat-square)](https://www.nuget.org/packages/XlsxToHtmlConverter/1.2.7)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg?style=flat-square)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.7/LICENSE.txt)

> A xlsx to html file converter. Support cell fill, font, border, alignment and other styles. Support custom column width and row height. Support vertical and/or horizontal merged cells. Support sheet tab color and hidden sheet. Support pictures drawing. Support progress callback event. It uses .Net 6.0 as framework and only depends on the Open Xml SDK.

## Dependencies

**.Net** >= 6.0  
**DocumentFormat.OpenXml** = 2.10.1

## Main Features

- [x] Cell fill, font, border, alignment, and other styles
- [x] Custom column width and row height
- [x] Vertical and/or horizontal merged cells
- [x] Sheet tab color and hidden sheet
- [x] Pictures drawing
- [x] Progress callback event

## How to Use

Only one line to convert a xlsx file to html string.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName);
```

Or if xlsx file data is in the stream, convert the stream.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileStream);
```

There are also flexible custom converter configurations.

```c#
XlsxToHtmlConverter.ConverterConfig config = new XlsxToHtmlConverter.ConverterConfig()
{
    PageTitle = "My Title",
    PresetStyles = "body { background-color: skyblue; } table { width: 100%; }",
    ErrorMessage = "Oh, no. It's error.",
    IsConvertStyles = true,
    IsConvertSizes = false,
    IsConvertPicture = true,
    IsConvertHiddenSheet = false
}

string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, config);
```

A progress callback event can be used to monitor the progress.

```c#
EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> converterProgressCallbackEvent = null;
converterProgressCallbackEvent += ConverterProgressCallback;

string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, converterProgressCallbackEvent);
```

And all of them can be used together.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, config, converterProgressCallbackEvent);
```

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.2.7/LICENSE.txt).
