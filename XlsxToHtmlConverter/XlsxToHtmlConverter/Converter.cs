using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxToHtmlConverter
{
    public class Converter
    {
        private const string alphabet = "abcdefghijklmnopqrstuvwxyz";

        private struct MergeCellInfo
        {
            public string FromColumnName { get; set; }
            public uint FromRowIndex { get; set; }
            public string ToColumnName { get; set; }
            public uint ToRowIndex { get; set; }
            public uint ColumnSpanned { get; set; }
            public uint RowSpanned { get; set; }
        }

        public static string ConvertXlsx(string fileName)
        {
            try
            {
                byte[] byteArray = File.ReadAllBytes(fileName);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(byteArray, 0, byteArray.Length);

                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(memoryStream, true))
                    {
                        string pageTitle = Path.GetFileName(fileName);

                        string tableHtml = "";

                        WorkbookPart workbook = doc.WorkbookPart;
                        WorkbookStylesPart styles = workbook.WorkbookStylesPart;
                        List<Sheet> sheets = workbook.Workbook.Descendants<Sheet>().ToList();
                        SharedStringTable sharedStringTable = workbook.GetPartsOfType<SharedStringTablePart>().FirstOrDefault().SharedStringTable;

                        int currentSheet = 0;

                        foreach (WorksheetPart worksheet in workbook.WorksheetParts)
                        {
                            tableHtml += $"\n{new string(' ', 4)}<h5>{sheets[currentSheet].Name}</h5>";
                            tableHtml += $"\n{new string(' ', 4)}<table>";

                            Worksheet sheet = worksheet.Worksheet;

                            bool containsMergeCells = false;

                            List<MergeCellInfo> mergeCells = new List<MergeCellInfo>();
                            if (sheet.Descendants<MergeCells>().FirstOrDefault() is MergeCells mergeCellsGroup)
                            {
                                containsMergeCells = true;

                                foreach (MergeCell mergeCell in mergeCellsGroup)
                                {
                                    try
                                    {
                                        string[] range = mergeCell.Reference.Value.Split(':');

                                        string firstColumn = GetColumnName(range[0]);
                                        string secondColumn = GetColumnName(range[1]);
                                        uint firstRow = GetRowIndex(range[0]);
                                        uint secondRow = GetRowIndex(range[1]);

                                        string fromColumn = alphabet.IndexOf(firstColumn.ToLower()) <= alphabet.IndexOf(secondColumn.ToLower()) ? firstColumn : secondColumn;
                                        string toColumn = fromColumn == firstColumn ? secondColumn : firstColumn;
                                        uint fromRow = Math.Min(firstRow, secondRow);
                                        uint toRow = Math.Max(firstRow, secondRow);

                                        mergeCells.Add(new MergeCellInfo() { FromColumnName = fromColumn, FromRowIndex = fromRow, ToColumnName = toColumn, ToRowIndex = toRow, ColumnSpanned = (uint)Math.Abs(alphabet.IndexOf(toColumn.ToLower()) - alphabet.IndexOf(fromColumn.ToLower())) + 1, RowSpanned = (uint)Math.Abs(toRow - fromRow) + 1 });
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                            List<Row> rows = sheet.Descendants<Row>().ToList();

                            uint totalRows = 0;
                            int totalColumn = 0;

                            foreach (Row row in rows)
                            {
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    try
                                    {
                                        string columnName = GetColumnName(cell.CellReference);
                                        uint rowIndex = GetRowIndex(cell.CellReference);

                                        int columnIndex = alphabet.IndexOf(columnName.ToLower()) + 1;

                                        if (totalColumn < columnIndex)
                                        {
                                            totalColumn = columnIndex;
                                        }
                                        if (totalRows < rowIndex)
                                        {
                                            totalRows = rowIndex;
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }

                            int currentColumn = 0;
                            int currentRow = 0;
                            uint lastRow = 0;

                            List<double> columnWidths = new List<double>();
                            for (uint i = 0; i < totalColumn; i++)
                            {
                                columnWidths.Add(Double.NaN);
                            }
                            if (sheet.GetFirstChild<Columns>() is Columns columnsGroup)
                            {
                                foreach (Column column in columnsGroup.Descendants<Column>())
                                {
                                    for (uint i = column.Min; i <= column.Max; i++)
                                    {
                                        try
                                        {
                                            if (column.CustomWidth == true && column.Width != null && column.Width.HasValue == true)
                                            {
                                                columnWidths[(int)i - 1] = CalculateColumnWidth(column.Width.Value);
                                            }
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                    }
                                }
                            }

                            foreach (Row row in rows)
                            {
                                if (row.RowIndex.Value - lastRow > 1)
                                {
                                    for (int i = 0; i < row.RowIndex.Value - lastRow - 1; i++)
                                    {
                                        tableHtml += $"\n{new string(' ', 8)}<tr>";

                                        for (int j = 0; j < totalColumn; j++)
                                        {
                                            double actualCellWidth = j >= columnWidths.Count ? Double.NaN : columnWidths[j];
                                            tableHtml += $"\n{new string(' ', 12)}<td style=\"height: 20px; width: {(Double.IsNaN(actualCellWidth) ? "auto" : actualCellWidth + "px")};\"></td>";
                                        }

                                        tableHtml += $"\n{new string(' ', 8)}</tr>";
                                    }
                                }

                                currentColumn = 0;

                                double rowHeight = 20;

                                if (row.CustomHeight != null && row.CustomHeight.Value == true && row.Height != null && row.Height.HasValue == true)
                                {
                                    rowHeight = row.Height.Value / 0.75;
                                }

                                tableHtml += $"\n{new string(' ', 8)}<tr>";

                                List<Cell> cells = new List<Cell>(totalColumn);

                                for (int i = 0; i < totalColumn; i++)
                                {
                                    cells.Add(new Cell() { CellValue = new CellValue(""), CellReference = alphabet[i].ToString().ToUpper() + row.RowIndex });
                                }
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    int actualCellIndex = alphabet.IndexOf(GetColumnName(cell.CellReference).ToLower());
                                    cells[actualCellIndex] = cell;
                                }

                                foreach (Cell cell in cells)
                                {
                                    int addedColumnNumber = 1;

                                    uint columnSpanned = 1;
                                    uint rowSpanned = 1;

                                    double actualCellHeight = rowHeight;
                                    double actualCellWidth = currentColumn >= columnWidths.Count ? Double.NaN : columnWidths[currentColumn];

                                    if (containsMergeCells == true && cell.CellReference != null)
                                    {
                                        string columnName = GetColumnName(cell.CellReference);
                                        uint rowIndex = GetRowIndex(cell.CellReference);

                                        if (mergeCells.Any(x => (rowIndex == x.FromRowIndex && columnName.ToLower() == x.FromColumnName.ToLower()) == false && rowIndex >= x.FromRowIndex && rowIndex <= x.ToRowIndex && alphabet.IndexOf(columnName.ToLower()) >= alphabet.IndexOf(x.FromColumnName.ToLower()) && alphabet.IndexOf(columnName.ToLower()) <= alphabet.IndexOf(x.ToColumnName.ToLower())) == true)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            foreach (MergeCellInfo mergeCellInfo in mergeCells)
                                            {
                                                if (columnName.ToLower() == mergeCellInfo.FromColumnName.ToLower() && rowIndex == mergeCellInfo.FromRowIndex)
                                                {
                                                    addedColumnNumber = (int)mergeCellInfo.ColumnSpanned;

                                                    columnSpanned = mergeCellInfo.ColumnSpanned;
                                                    rowSpanned = mergeCellInfo.RowSpanned;

                                                    if (rowSpanned > 1)
                                                    {
                                                        actualCellHeight = 0;

                                                        for (int i = 0; i < rowSpanned; i++)
                                                        {
                                                            int index = currentRow + i;
                                                            double height = 20;

                                                            if (rows[index].CustomHeight != null && rows[index].CustomHeight.Value == true && row.Height != null && row.Height.HasValue == true)
                                                            {
                                                                height = rows[index].Height.Value / 0.75;
                                                            }

                                                            actualCellHeight += height;
                                                        }

                                                        actualCellHeight -= rowSpanned - 1;
                                                    }

                                                    break;
                                                }
                                            }
                                        }
                                    }

                                    string cellValue = "";

                                    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                                    {
                                        int ssid = int.Parse(cell.CellValue.Text);
                                        SharedStringItem sharedString = (SharedStringItem)sharedStringTable.ChildElements[ssid];

                                        try
                                        {
                                            foreach (OpenXmlElement element in sharedString.Descendants())
                                            {
                                                if (element is Text text)
                                                {
                                                    cellValue += text.Text;
                                                }
                                                else if (element is Run run)
                                                {
                                                    string runStyle = "";

                                                    if (run.RunProperties.GetFirstChild<Color>() is Color fontColor)
                                                    {
                                                        string value = GetColorFromColorType(doc, fontColor, System.Drawing.Color.Black);

                                                        runStyle += $" color: {value};";
                                                    }
                                                    if (run.RunProperties.GetFirstChild<Bold>() is Bold bold)
                                                    {
                                                        bool value = bold.Val != null ? bold.Val.Value : false;

                                                        runStyle += $" font-weight: {(value ? "bold" : "normal")};";
                                                    }
                                                    if (run.RunProperties.GetFirstChild<Italic>() is Italic italic)
                                                    {
                                                        bool value = italic.Val != null ? italic.Val.Value : false;

                                                        runStyle += $" font-style: {(value ? "italic" : "normal")};";
                                                    }
                                                    if (run.RunProperties.GetFirstChild<Strike>() is Strike strike)
                                                    {
                                                        bool value = strike.Val != null ? strike.Val.Value : false;

                                                        runStyle += $" text-decoration: {(value ? "line-through" : "none")}";

                                                        if (run.RunProperties.GetFirstChild<Underline>() is Underline underline)
                                                        {
                                                            string underlineValue = underline.Val != null ? underline.Val.Value switch
                                                            {
                                                                UnderlineValues.Single => " underline;",
                                                                UnderlineValues.SingleAccounting => " underline;",
                                                                UnderlineValues.Double => " underline;",
                                                                UnderlineValues.DoubleAccounting => " underline;",
                                                                _ => ";",
                                                            } : ";";
                                                            runStyle += underlineValue;
                                                        }
                                                        else
                                                        {
                                                            runStyle += ";";
                                                        }
                                                    }
                                                    if (run.RunProperties.GetFirstChild<FontSize>() is FontSize fontSize)
                                                    {
                                                        double value = fontSize.Val != null ? fontSize.Val.Value * 96 / 72 : 11;

                                                        runStyle += $" font-size: {fontSize}px;";
                                                    }
                                                    if (run.RunProperties.GetFirstChild<RunFont>() is RunFont runFont)
                                                    {
                                                        string value = runFont.Val != null ? runFont.Val.Value : "serif";

                                                        runStyle += $" font-family: {value};";
                                                    }

                                                    cellValue += $"<p style=\"{runStyle}\">{run.Text}</p>";
                                                }
                                                else
                                                {
                                                    cellValue += element.InnerText;
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            cellValue = sharedString.InnerText;
                                        }
                                    }
                                    else if (cell.CellValue != null)
                                    {
                                        cellValue = cell.CellValue.Text;
                                    }
                                    else if (cell.InnerText != null)
                                    {
                                        cellValue = cell.InnerText;
                                    }

                                    string advancedStyleHtml = $"";

                                    if (cell.StyleIndex != null)
                                    {
                                        if (styles.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value] is CellFormat cellFormat)
                                        {
                                            try
                                            {
                                                if (styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value] is Fill fill)
                                                {
                                                    string background = fill.PatternFill != null && fill.PatternFill.PatternType.Value != PatternValues.None && fill.PatternFill.ForegroundColor != null ? GetColorFromColorType(doc, fill.PatternFill.ForegroundColor, System.Drawing.Color.Transparent) : "transparent";

                                                    advancedStyleHtml += $" background-color: {background};";
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Diagnostics.Debug.Write("Get fill style failed. Error: " + ex.Message);
                                            }
                                            try
                                            {
                                                if (styles.Stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value] is Font font)
                                                {
                                                    string fontColor = font.Color != null ? GetColorFromColorType(doc, font.Color, System.Drawing.Color.Black) : "black";
                                                    double fontSize = font.FontSize != null && font.FontSize.Val != null ? font.FontSize.Val.Value * 96 / 72 : 11;
                                                    string fontFamily = font.FontName != null && font.FontName.Val != null ? $"\'{font.FontName.Val.Value}\', serif" : "serif";
                                                    bool bold = font.Bold != null && font.Bold.Val != null ? font.Bold.Val.Value : false;
                                                    bool italic = font.Italic != null && font.Italic.Val != null ? font.Italic.Val.Value : false;
                                                    bool strike = font.Strike != null && font.Strike.Val != null ? font.Strike.Val.Value : false;
                                                    string underline = font.Underline != null && font.Underline.Val != null ? font.Underline.Val.Value switch
                                                    {
                                                        UnderlineValues.Single => " underline",
                                                        UnderlineValues.SingleAccounting => " underline",
                                                        UnderlineValues.Double => " underline",
                                                        UnderlineValues.DoubleAccounting => " underline",
                                                        _ => "",
                                                    } : "";

                                                    advancedStyleHtml += $" color: {fontColor}; font-size: {fontSize}px; font-family: {fontFamily}; font-weight: {(bold ? "bold" : "normal")}; font-style: {(italic ? "italic" : "normal")}; text-decoration: {(strike ? "line-through" : "none")}{underline};";
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Diagnostics.Debug.Write("Get font style failed. Error: " + ex.Message);
                                            }
                                            try
                                            {
                                                if (styles.Stylesheet.Borders.ChildElements[(int)cellFormat.BorderId.Value] is Border border)
                                                {
                                                    string leftWidth = "1px";
                                                    string rightWidth = "1px";
                                                    string topWidth = "1px";
                                                    string bottomWidth = "1px";

                                                    string leftStyle = "solid";
                                                    string rightStyle = "solid";
                                                    string topStyle = "solid";
                                                    string bottomStyle = "solid";

                                                    string leftColor = "lightgray";
                                                    string rightColor = "lightgray";
                                                    string topColor = "lightgray";
                                                    string bottomColor = "lightgray";

                                                    if (border.LeftBorder is LeftBorder leftBorder)
                                                    {
                                                        BorderStyleToHtmlAttributes(doc, leftBorder, ref leftWidth, ref leftStyle, ref leftColor);
                                                    }
                                                    if (border.RightBorder is RightBorder rightBorder)
                                                    {
                                                        BorderStyleToHtmlAttributes(doc, rightBorder, ref rightWidth, ref rightStyle, ref rightColor);
                                                    }
                                                    if (border.TopBorder is TopBorder topBorder)
                                                    {
                                                        BorderStyleToHtmlAttributes(doc, topBorder, ref topWidth, ref topStyle, ref topColor);
                                                    }
                                                    if (border.BottomBorder is BottomBorder bottomBorder)
                                                    {
                                                        BorderStyleToHtmlAttributes(doc, bottomBorder, ref bottomWidth, ref bottomStyle, ref bottomColor);
                                                    }

                                                    advancedStyleHtml += $" border-width: {topWidth} {rightWidth} {bottomWidth} {leftWidth}; border-style: {topStyle} {rightStyle} {bottomStyle} {leftStyle}; border-color: {topColor} {rightColor} {bottomColor} {leftColor};";
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Diagnostics.Debug.Write("Get font style failed. Error: " + ex.Message);
                                            }
                                            try
                                            {
                                                if (cellFormat.Alignment != null)
                                                {
                                                    string verticalTextAlignment = "bottom";
                                                    string horizontalTextAlignment = "left";

                                                    if (cellFormat.Alignment.Vertical != null)
                                                    {
                                                        verticalTextAlignment = cellFormat.Alignment.Vertical.Value switch
                                                        {
                                                            VerticalAlignmentValues.Bottom => "bottom",
                                                            VerticalAlignmentValues.Center => "middle",
                                                            VerticalAlignmentValues.Top => "top",
                                                            _ => "bottom",
                                                        };
                                                    }
                                                    if (cellFormat.Alignment.Horizontal != null)
                                                    {
                                                        horizontalTextAlignment = cellFormat.Alignment.Horizontal.Value switch
                                                        {
                                                            HorizontalAlignmentValues.Left => "left",
                                                            HorizontalAlignmentValues.Center => "center",
                                                            HorizontalAlignmentValues.Right => "right",
                                                            HorizontalAlignmentValues.Justify => "justify",
                                                            _ => "bottom",
                                                        };
                                                    }

                                                    advancedStyleHtml += $" text-align: {horizontalTextAlignment}; vertical-align: {verticalTextAlignment};";

                                                    if (cellFormat.Alignment.WrapText != null && cellFormat.Alignment.WrapText.Value == true)
                                                    {
                                                        advancedStyleHtml += " word-wrap: break-word; white-space: normal;";
                                                    }
                                                    if (cellFormat.Alignment.TextRotation != null)
                                                    {
                                                        cellValue = $"<div style=\"width: min-content; transform: rotate(-{cellFormat.Alignment.TextRotation.Value}deg);\">" + cellValue + "</div>";
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Diagnostics.Debug.Write("Get alignment style failed. Error: " + ex.Message);
                                            }
                                        }
                                    }

                                    tableHtml += $"\n{new string(' ', 12)}<td colspan=\"{columnSpanned}\" rowspan=\"{rowSpanned}\" style=\"height: {(Double.IsNaN(actualCellHeight) ? "auto" : actualCellHeight + "px")}; width: {(Double.IsNaN(actualCellWidth) ? "auto" : actualCellWidth + "px")};{advancedStyleHtml}\">{cellValue}</td>";

                                    currentColumn += addedColumnNumber;
                                }

                                tableHtml += $"\n{new string(' ', 8)}</tr>";

                                currentRow++;
                                lastRow = row.RowIndex.Value;
                            }

                            tableHtml += $"\n{new string(' ', 4)}</table>";

                            currentSheet++;
                        }

                        string htmlString = String.Format(@"<!DOCTYPE html>
<html lang=""en"">

<head>
    <meta charset=""UTF-8"">
    <title>{0}</title>

    <style>
        body {{
            margin: 0;
            padding: 0;
            width: 100%;
        }}

        h5 {{
            font-size: 20px;
            font-weight: bold;
            text-align: center;
            margin: 10px auto;
        }}

        table {{
            width: 100%;
            table-layout：fixed;
            border-collapse: collapse;
        }}

        td {{
            text-align: left;
            vertical-align: bottom;
            padding: 2px;
            color: black;
            background-color: transparent;
            border-width: 1px;
            border-style: solid;
            border-color: lightgray;
            border-collapse: collapse;
            white-space: nowrap;
        }}
    </style>
</head>
<body>
    {1}
</body>
</html>", pageTitle, tableHtml);

                        //File.WriteAllText("D:\\test.html", htmlString);

                        return htmlString;
                    }
                }
            }
            catch
            {
                return "Error, can not convert XLSX file. The file is either already open (please close it) or contains corrupt data.";
            }
        }

        private static string GetColumnName(string cellName)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        private static uint GetRowIndex(string cellName)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        private static void BorderStyleToHtmlAttributes(SpreadsheetDocument doc, BorderPropertiesType border, ref string width, ref string style, ref string color)
        {
            if (border.Style != null)
            {
                switch (border.Style.Value)
                {
                    case BorderStyleValues.None:
                        width = "0";
                        style = "solid";
                        break;
                    case BorderStyleValues.Thin:
                        width = "thin";
                        style = "solid";
                        break;
                    case BorderStyleValues.Medium:
                        width = "medium";
                        style = "solid";
                        break;
                    case BorderStyleValues.MediumDashDot:
                        width = "medium";
                        style = "dashed";
                        break;
                    case BorderStyleValues.MediumDashDotDot:
                        width = "medium";
                        style = "dotted";
                        break;
                    case BorderStyleValues.MediumDashed:
                        width = "medium";
                        style = "solid";
                        break;
                    case BorderStyleValues.Thick:
                        width = "thick";
                        style = "solid";
                        break;
                    case BorderStyleValues.Dashed:
                        width = "1px";
                        style = "dashed";
                        break;
                    case BorderStyleValues.DashDot:
                        width = "1px";
                        style = "dashed";
                        break;
                    case BorderStyleValues.DashDotDot:
                        width = "1px";
                        style = "dashed";
                        break;
                    case BorderStyleValues.Dotted:
                        width = "1px";
                        style = "dotted";
                        break;
                    case BorderStyleValues.Double:
                        width = "1px";
                        style = "double";
                        break;
                    case BorderStyleValues.Hair:
                        width = "1px";
                        style = "solid";
                        break;
                    case BorderStyleValues.SlantDashDot:
                        width = "1px";
                        style = "dashed";
                        break;
                }
            }

            color = border.Color != null ? GetColorFromColorType(doc, border.Color, System.Drawing.Color.LightGray) : "lightgray";
        }

        private static double CalculateColumnWidth(double textLength)
        {
            return (textLength - 1) * 7 - 5 + 12;
        }

        private static string GetColorFromColorType(SpreadsheetDocument doc, ColorType type, System.Drawing.Color defaultColor)
        {
            System.Drawing.Color rgbColor;

            try
            {
                if (type == null)
                {
                    throw new Exception("Doesn't have any color. Please use default value.");
                }

                if (type.Rgb != null)
                {
                    rgbColor = System.Drawing.ColorTranslator.FromHtml("#" + type.Rgb.Value);
                }
                else if (type.Indexed != null)
                {
                    rgbColor = System.Drawing.ColorTranslator.FromHtml(IndexedColorsList.GetIndexColour(type.Indexed.Value));
                }
                else if (type.Theme != null)
                {
                    DocumentFormat.OpenXml.Drawing.Color2Type color = (DocumentFormat.OpenXml.Drawing.Color2Type)doc.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)type.Theme.Value];

                    if (color.RgbColorModelHex != null)
                    {
                        rgbColor = System.Drawing.ColorTranslator.FromHtml(color.RgbColorModelHex.Val.Value);
                    }
                    else if (color.RgbColorModelPercentage != null)
                    {
                        rgbColor = System.Drawing.Color.FromArgb(color.RgbColorModelPercentage.RedPortion / 1000 / 100 * 255, color.RgbColorModelPercentage.GreenPortion / 1000 / 100 * 255, color.RgbColorModelPercentage.BluePortion / 1000 / 100 * 255);
                    }
                    else if (color.HslColor != null)
                    {
                        HlsToRgb(color.HslColor.HueValue, color.HslColor.LumValue, color.HslColor.SatValue, out int r, out int g, out int b);
                        
                        rgbColor = System.Drawing.Color.FromArgb(r, g, b);
                    }
                    else if (color.SystemColor != null)
                    {
                        rgbColor = color.SystemColor.Val.Value switch
                        {
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveBorder => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ActiveBorder),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveCaption => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ActiveCaption),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ApplicationWorkspace => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.AppWorkspace),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.Background => System.Drawing.Color.Transparent,
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonFace => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Control),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonHighlight => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Highlight),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonShadow => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ControlDark),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ControlText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.CaptionText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ActiveCaptionText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientActiveCaption => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.GradientActiveCaption),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientInactiveCaption => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.GradientInactiveCaption),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.GrayText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.GrayText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.Highlight => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Highlight),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.HighlightText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.HighlightText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.HotLight => System.Drawing.Color.Orange,
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveBorder => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.InactiveBorder),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaption => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.InactiveCaption),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaptionText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.InactiveCaptionText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoBack => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Info),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.InfoText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.Menu => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Menu),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuBar => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.MenuBar),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuHighlight => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.MenuHighlight),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.MenuText),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ScrollBar => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ScrollBar),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDDarkShadow => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ControlDark),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDLight => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ControlLight),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.Window => System.Drawing.Color.Black,
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowFrame => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.WindowFrame),
                            DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText => System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.WindowText),
                            _ => throw new Exception("Doesn't have any custom value. Please use default value.")
                        };
                    }
                    else if (color.PresetColor != null)
                    {
                        rgbColor = color.PresetColor.Val.Value switch
                        {
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.AliceBlue => System.Drawing.Color.AliceBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.AntiqueWhite => System.Drawing.Color.AntiqueWhite,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Aqua => System.Drawing.Color.Aqua,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Aquamarine => System.Drawing.Color.Aquamarine,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Azure => System.Drawing.Color.Azure,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Beige => System.Drawing.Color.Beige,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Bisque => System.Drawing.Color.Bisque,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Black => System.Drawing.Color.Black,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.BlanchedAlmond => System.Drawing.Color.BlanchedAlmond,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Blue => System.Drawing.Color.Blue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.BlueViolet => System.Drawing.Color.BlueViolet,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Brown => System.Drawing.Color.Brown,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.BurlyWood => System.Drawing.Color.BurlyWood,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.CadetBlue => System.Drawing.Color.CadetBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Chartreuse => System.Drawing.Color.Chartreuse,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Chocolate => System.Drawing.Color.Chocolate,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Coral => System.Drawing.Color.Coral,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.CornflowerBlue => System.Drawing.Color.CornflowerBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Cornsilk => System.Drawing.Color.Cornsilk,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Crimson => System.Drawing.Color.Crimson,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Cyan => System.Drawing.Color.Cyan,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue => System.Drawing.Color.DarkBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan => System.Drawing.Color.DarkCyan,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod => System.Drawing.Color.DarkGoldenrod,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray => System.Drawing.Color.DarkGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen => System.Drawing.Color.DarkGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki => System.Drawing.Color.DarkKhaki,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta => System.Drawing.Color.DarkMagenta,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen => System.Drawing.Color.DarkOliveGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange => System.Drawing.Color.DarkOrange,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid => System.Drawing.Color.DarkOrchid,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed => System.Drawing.Color.DarkRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon => System.Drawing.Color.DarkSalmon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen => System.Drawing.Color.DarkSeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue => System.Drawing.Color.DarkSlateBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray => System.Drawing.Color.DarkSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise => System.Drawing.Color.DarkTurquoise,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet => System.Drawing.Color.DarkViolet,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepPink => System.Drawing.Color.DeepPink,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepSkyBlue => System.Drawing.Color.DeepSkyBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGray => System.Drawing.Color.DimGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DodgerBlue => System.Drawing.Color.DodgerBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Firebrick => System.Drawing.Color.Firebrick,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.FloralWhite => System.Drawing.Color.FloralWhite,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.ForestGreen => System.Drawing.Color.ForestGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Fuchsia => System.Drawing.Color.Fuchsia,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Gainsboro => System.Drawing.Color.Gainsboro,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.GhostWhite => System.Drawing.Color.GhostWhite,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Gold => System.Drawing.Color.Gold,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Goldenrod => System.Drawing.Color.Goldenrod,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Gray => System.Drawing.Color.Gray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Green => System.Drawing.Color.Green,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.GreenYellow => System.Drawing.Color.GreenYellow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Honeydew => System.Drawing.Color.Honeydew,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.HotPink => System.Drawing.Color.HotPink,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.IndianRed => System.Drawing.Color.IndianRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Indigo => System.Drawing.Color.Indigo,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Ivory => System.Drawing.Color.Ivory,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Khaki => System.Drawing.Color.Khaki,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Lavender => System.Drawing.Color.Lavender,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LavenderBlush => System.Drawing.Color.LavenderBlush,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LawnGreen => System.Drawing.Color.LawnGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LemonChiffon => System.Drawing.Color.LemonChiffon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue => System.Drawing.Color.LightBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral => System.Drawing.Color.LightCoral,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan => System.Drawing.Color.LightCyan,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow => System.Drawing.Color.LightGoldenrodYellow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray => System.Drawing.Color.LightGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen => System.Drawing.Color.LightGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink => System.Drawing.Color.LightPink,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon => System.Drawing.Color.LightSalmon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen => System.Drawing.Color.LightSeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue => System.Drawing.Color.LightSkyBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray => System.Drawing.Color.LightSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue => System.Drawing.Color.LightSteelBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow => System.Drawing.Color.LightYellow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Lime => System.Drawing.Color.Lime,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LimeGreen => System.Drawing.Color.LimeGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Linen => System.Drawing.Color.Linen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Magenta => System.Drawing.Color.Magenta,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Maroon => System.Drawing.Color.Maroon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MedAquamarine => System.Drawing.Color.MediumAquamarine,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue => System.Drawing.Color.MediumBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid => System.Drawing.Color.MediumOrchid,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple => System.Drawing.Color.MediumPurple,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen => System.Drawing.Color.MediumSeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue => System.Drawing.Color.MediumSlateBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen => System.Drawing.Color.MediumSpringGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise => System.Drawing.Color.MediumTurquoise,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed => System.Drawing.Color.MediumVioletRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MidnightBlue => System.Drawing.Color.MidnightBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MintCream => System.Drawing.Color.MintCream,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MistyRose => System.Drawing.Color.MistyRose,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Moccasin => System.Drawing.Color.Moccasin,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.NavajoWhite => System.Drawing.Color.NavajoWhite,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Navy => System.Drawing.Color.Navy,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.OldLace => System.Drawing.Color.OldLace,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Olive => System.Drawing.Color.Olive,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.OliveDrab => System.Drawing.Color.OliveDrab,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Orange => System.Drawing.Color.Orange,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.OrangeRed => System.Drawing.Color.OrangeRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Orchid => System.Drawing.Color.Orchid,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGoldenrod => System.Drawing.Color.PaleGoldenrod,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGreen => System.Drawing.Color.PaleGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleTurquoise => System.Drawing.Color.PaleTurquoise,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleVioletRed => System.Drawing.Color.PaleVioletRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PapayaWhip => System.Drawing.Color.PapayaWhip,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PeachPuff => System.Drawing.Color.PeachPuff,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Peru => System.Drawing.Color.Peru,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Pink => System.Drawing.Color.Pink,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Plum => System.Drawing.Color.Plum,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.PowderBlue => System.Drawing.Color.PowderBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Purple => System.Drawing.Color.Purple,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Red => System.Drawing.Color.Red,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.RosyBrown => System.Drawing.Color.RosyBrown,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.RoyalBlue => System.Drawing.Color.RoyalBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SaddleBrown => System.Drawing.Color.SaddleBrown,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Salmon => System.Drawing.Color.Salmon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SandyBrown => System.Drawing.Color.SandyBrown,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaGreen => System.Drawing.Color.SeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaShell => System.Drawing.Color.SeaShell,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Sienna => System.Drawing.Color.Sienna,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Silver => System.Drawing.Color.Silver,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SkyBlue => System.Drawing.Color.SkyBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateBlue => System.Drawing.Color.SlateBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGray => System.Drawing.Color.SlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Snow => System.Drawing.Color.Snow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SpringGreen => System.Drawing.Color.SpringGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SteelBlue => System.Drawing.Color.SteelBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Tan => System.Drawing.Color.Tan,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Teal => System.Drawing.Color.Teal,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Thistle => System.Drawing.Color.Thistle,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Tomato => System.Drawing.Color.Tomato,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Turquoise => System.Drawing.Color.Turquoise,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Violet => System.Drawing.Color.Violet,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Wheat => System.Drawing.Color.Wheat,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.White => System.Drawing.Color.White,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.WhiteSmoke => System.Drawing.Color.WhiteSmoke,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Yellow => System.Drawing.Color.Yellow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.YellowGreen => System.Drawing.Color.YellowGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue2010 => System.Drawing.Color.DarkBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan2010 => System.Drawing.Color.DarkCyan,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod2010 => System.Drawing.Color.DarkGoldenrod,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray2010 => System.Drawing.Color.DarkGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey2010 => System.Drawing.Color.DarkGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen2010 => System.Drawing.Color.DarkGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki2010 => System.Drawing.Color.DarkKhaki,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta2010 => System.Drawing.Color.DarkMagenta,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen2010 => System.Drawing.Color.DarkOliveGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange2010 => System.Drawing.Color.DarkOrange,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid2010 => System.Drawing.Color.DarkOrchid,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed2010 => System.Drawing.Color.DarkRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon2010 => System.Drawing.Color.DarkSalmon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen2010 => System.Drawing.Color.DarkSeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue2010 => System.Drawing.Color.DarkSlateBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray2010 => System.Drawing.Color.DarkSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey2010 => System.Drawing.Color.DarkSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise2010 => System.Drawing.Color.DarkTurquoise,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet2010 => System.Drawing.Color.DarkViolet,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue2010 => System.Drawing.Color.LightBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral2010 => System.Drawing.Color.LightCoral,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan2010 => System.Drawing.Color.LightCyan,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow2010 => System.Drawing.Color.LightGoldenrodYellow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray2010 => System.Drawing.Color.LightGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey2010 => System.Drawing.Color.LightGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen2010 => System.Drawing.Color.LightGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink2010 => System.Drawing.Color.LightPink,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon2010 => System.Drawing.Color.LightSalmon,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen2010 => System.Drawing.Color.LightSeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue2010 => System.Drawing.Color.LightSkyBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray2010 => System.Drawing.Color.LightSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey2010 => System.Drawing.Color.LightSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue2010 => System.Drawing.Color.LightSteelBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow2010 => System.Drawing.Color.LightYellow,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumAquamarine2010 => System.Drawing.Color.MediumAquamarine,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue2010 => System.Drawing.Color.MediumBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid2010 => System.Drawing.Color.MediumOrchid,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple2010 => System.Drawing.Color.MediumPurple,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen2010 => System.Drawing.Color.MediumSeaGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue2010 => System.Drawing.Color.MediumSlateBlue,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen2010 => System.Drawing.Color.MediumSpringGreen,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise2010 => System.Drawing.Color.MediumTurquoise,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed2010 => System.Drawing.Color.MediumVioletRed,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey => System.Drawing.Color.DarkGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGrey => System.Drawing.Color.DimGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey => System.Drawing.Color.DarkSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.Grey => System.Drawing.Color.Gray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey => System.Drawing.Color.LightGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey => System.Drawing.Color.LightSlateGray,
                            DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGrey => System.Drawing.Color.SlateGray,
                            _ => throw new Exception("Doesn't have any custom value. Please use default value."),
                        };
                    }
                    else
                    {
                        throw new Exception("Doesn't have any custom value. Please use default value.");
                    }
                }
                else
                {
                    throw new Exception("Doesn't have any custom value. Please use default value.");
                }
            }
            catch
            {
                rgbColor = defaultColor;
            }

            try
            {
                if (type != null && type.Tint != null)
                {
                    double tint = type.Tint.Value;

                    if (tint < 0)
                    {
                        RgbToHls(rgbColor.R, rgbColor.G, rgbColor.B, out double h, out double l, out double s);

                        l *= (1.0 + tint);

                        HlsToRgb(h, l, s, out int r, out int g, out int b);

                        rgbColor = System.Drawing.Color.FromArgb(r, g, b);
                    }
                    else if (type.Tint.Value < 0)
                    {
                        RgbToHls(rgbColor.R, rgbColor.G, rgbColor.B, out double h, out double l, out double s);

                        double max = h;
                        if (max < l)
                        {
                            max = l;
                        }
                        if (max < s)
                        {
                            max = s;
                        }

                        l *= 1.0 - tint;
                        l += max - max * (1.0 - tint);

                        HlsToRgb(h, l, s, out int r, out int g, out int b);

                        rgbColor = System.Drawing.Color.FromArgb(r, g, b);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write("Get color tint failed. Error: " + ex.Message);
            }

            return $"rgb({rgbColor.R}, {rgbColor.G}, {rgbColor.B})";
        }

        private static void RgbToHls(int r, int g, int b, out double h, out double l, out double s)
        {
            double double_r = r / 255.0;
            double double_g = g / 255.0;
            double double_b = b / 255.0;

            double max = double_r;
            if (max < double_g)
            {
                max = double_g;
            }
            if (max < double_b)
            {
                max = double_b;
            }

            double min = double_r;
            if (min > double_g)
            {
                min = double_g;
            }
            if (min > double_b)
            {
                min = double_b;
            }

            double diff = max - min;
            l = (max + min) / 2;
            if (Math.Abs(diff) < 0.00001)
            {
                s = 0;
                h = 0;
            }
            else
            {
                if (l <= 0.5)
                {
                    s = diff / (max + min);
                }
                else
                {
                    s = diff / (2 - max - min);
                }

                double r_dist = (max - double_r) / diff;
                double g_dist = (max - double_g) / diff;
                double b_dist = (max - double_b) / diff;

                if (double_r == max)
                {
                    h = b_dist - g_dist;
                }
                else if (double_g == max)
                {
                    h = 2 + r_dist - b_dist;
                }
                else
                {
                    h = 4 + g_dist - r_dist;
                }

                h *= 60;
                if (h < 0)
                {
                    h += 360;
                }
            }
        }

        private static void HlsToRgb(double h, double l, double s, out int r, out int g, out int b)
        {
            double p2;
            if (l <= 0.5)
            {
                p2 = l * (1 + s);
            }
            else
            {
                p2 = l + s - l * s;
            }

            double p1 = 2 * l - p2;
            double double_r, double_g, double_b;
            if (s == 0)
            {
                double_r = l;
                double_g = l;
                double_b = l;
            }
            else
            {
                double_r = QqhToRgb(p1, p2, h + 120);
                double_g = QqhToRgb(p1, p2, h);
                double_b = QqhToRgb(p1, p2, h - 120);
            }

            r = (int)(double_r * 255.0);
            g = (int)(double_g * 255.0);
            b = (int)(double_b * 255.0);
        }

        private static double QqhToRgb(double q1, double q2, double hue)
        {
            if (hue > 360)
            {
                hue -= 360;
            }
            else if (hue < 0)
            {
                hue += 360;
            }

            if (hue < 60)
            {
                return q1 + (q2 - q1) * hue / 60;
            }
            else if (hue < 180)
            {
                return q2;
            }
            else if (hue < 240)
            {
                return q1 + (q2 - q1) * (240 - hue) / 60;
            }
            return q1;
        }

        private class IndexedColorsList
        {
            public static readonly Dictionary<uint, string> Data = new Dictionary<uint, string>()
            {
                { 0, "#000000" },
                { 1, "#FFFFFF" },
                { 2, "#FF0000" },
                { 3, "#00FF00" },
                { 4, "#0000FF" },
                { 5, "#FFFF00" },
                { 6, "#FF00FF" },
                { 7, "#00FFFF" },
                { 8, "#000000" },
                { 9, "#FFFFFF" },
                { 10, "#FF0000" },
                { 11, "#00FF00" },
                { 12, "#0000FF" },
                { 13, "#FFFF00" },
                { 14, "#FF00FF" },
                { 15, "#00FFFF" },
                { 16, "#800000" },
                { 17, "#008000" },
                { 18, "#000080" },
                { 19, "#808000" },
                { 20, "#800080" },
                { 21, "#008080" },
                { 22, "#C0C0C0" },
                { 23, "#808080" },
                { 24, "#9999FF" },
                { 25, "#993366" },
                { 26, "#FFFFCC" },
                { 27, "#CCFFFF" },
                { 28, "#660066" },
                { 29, "#FF8080" },
                { 30, "#0066CC" },
                { 31, "#CCCCFF" },
                { 32, "#000080" },
                { 33, "#FF00FF" },
                { 34, "#FFFF00" },
                { 35, "#00FFFF" },
                { 36, "#800080" },
                { 37, "#800000" },
                { 38, "#008080" },
                { 39, "#0000FF" },
                { 40, "#00CCFF" },
                { 41, "#CCFFFF" },
                { 42, "#CCFFCC" },
                { 43, "#FFFF99" },
                { 44, "#99CCFF" },
                { 45, "#FF99CC" },
                { 46, "#CC99FF" },
                { 47, "#FFCC99" },
                { 48, "#3366FF" },
                { 49, "#33CCCC" },
                { 50, "#99CC00" },
                { 51, "#FFCC00" },
                { 52, "#FF9900" },
                { 53, "#FF6600" },
                { 54, "#666699" },
                { 55, "#969696" },
                { 56, "#003366" },
                { 57, "#339966" },
                { 58, "#003300" },
                { 59, "#333300" },
                { 60, "#993300" },
                { 61, "#993366" },
                { 62, "#333399" },
                { 63, "#333333" }
            };

            public static string GetIndexColour(uint index)
            {
                if (Data.TryGetValue(index, out string color) == true)
                {
                    return color;
                }
                else
                {
                    throw new Exception("Doesn't have any color value from index. Please use default value.");
                }
            }
        }
    }
}
