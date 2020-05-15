# XlsxToHtmlConverter

A xlsx to html file converter. Support fill, font, border, alignment and other styles. Support custom column width and row height. Support vertical and/or horizontal merged cells. It uses .Net Core 3.0 as framework and only depends on the Open Xml SDK.

[Click here to view the latest version](https://github.com/Jet20070731/XlsxToHtmlConverter/tree/1.0.3)

## Dependencies

**.Net Core** >= 3.0  
**DocumentFormat.OpenXml** = 2.10.1

## How to Use

Only one line to convert xlsx file to html string.

```c#
string html = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName);
```

## Commercial Samples

#### Otpos PDF Editor

[![Fivicon](http://pdf-editor.otpos.com/content/img/favicon.png)](http://pdf-editor.otpos.com/)

## License

This project is under the [MIT License](https://github.com/Jet20070731/XlsxToHtmlConverter/blob/1.0.3/LICENSE.txt).
