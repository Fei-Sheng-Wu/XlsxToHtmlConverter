# XlsxToHtmlConverter

[![Target Framework](https://img.shields.io/badge/%2ENet%20Core-3.0-green.svg)](https://docs.microsoft.com/en-us/dotnet/core/about)
[![Nuget](https://img.shields.io/badge/Nuget-v1.1.8-blue.svg)](https://docs.microsoft.com/en-us/dotnet/core/about)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg)](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.1.8/LICENSE.txt)

> A xlsx to html file converter. Support fill, font, border, alignment and other styles. Support custom column width and row height. Support vertical and/or horizontal merged cells. Support sheet tab color and hidden sheet. Support pictures drawing. Support progress callback event. It uses .Net Core 3.0 as framework and only depends on the Open Xml SDK.

## Dependencies

**.Net Core** >= 3.0  
**DocumentFormat.OpenXml** = 2.10.1

## Main Features

- [x] Cell fill, font, border, alignment, and other styles
- [x] Custom column width and row height
- [x] Vertical and/or horizontal merged cells
- [x] Sheet tab color and hidden sheet
- [x] Pictures drawing
- [x] Progress callback event

## How to Use

Only one line to convert xlsx file to html string.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName);
```

Or if xlsx file data is in the stream, convert the stream.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileStream);
```

You can even set your custom converter config.

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

And you can convert file with progress callback event.

```c#
EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> converterProgressCallbackEvent = null;
converterProgressCallbackEvent += ConverterProgressCallback;

string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, converterProgressCallbackEvent);
```

Also, you can use custom config and progress callback event together.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, config, converterProgressCallbackEvent);
```

## Commercial Samples

|Otpos PDF Editor|
|    :--------:   |
|[![Fivicon](http://pdf-editor.otpos.com/content/img/favicon.png)](http://pdf-editor.otpos.com/)|

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter/blob/1.1.8/LICENSE.txt).
