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
    /// <summary>
    /// The Xlsx to Html converter class.
    /// </summary>
    public class Converter
    {
        protected Converter()
        {
            return;
        }

        #region Public Methods

        /// <summary>
        /// Convert a local Xlsx file to Html string.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed or not.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, bool loadIntoMemory = true)
        {
            ConvertXlsx(fileName, outputHtml, ConverterConfig.DefaultSettings, null, loadIntoMemory);
        }

        /// <summary>
        /// Convert a local Xlsx file to Html string with specific configuartions.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed or not.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, ConverterConfig config, bool loadIntoMemory = true)
        {
            ConvertXlsx(fileName, outputHtml, config, null, loadIntoMemory);
        }

        /// <summary>
        /// Convert a local Xlsx file to Html string with progress callback event.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed or not.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, EventHandler<ConverterProgressCallbackEventArgs> progressCallback, bool loadIntoMemory = true)
        {
            ConvertXlsx(fileName, outputHtml, ConverterConfig.DefaultSettings, progressCallback, loadIntoMemory);
        }

        /// <summary>
        /// Convert a local Xlsx file to Html string with specific configuartions and progress callback event.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed or not.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, ConverterConfig config, EventHandler<ConverterProgressCallbackEventArgs> progressCallback, bool loadIntoMemory = true)
        {
            if (loadIntoMemory)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    byte[] byteArray = File.ReadAllBytes(fileName);
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    ConvertXlsx(memoryStream, outputHtml, config, progressCallback);
                }
            }
            else
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
                {
                    ConvertXlsx(fileStream, outputHtml, config, progressCallback);
                }
            }
        }

        /// <summary>
        /// Convert a stream Xlsx file to Html string.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml)
        {
            ConvertXlsx(inputXlsx, outputHtml, ConverterConfig.DefaultSettings, null);
        }

        /// <summary>
        /// Convert a stream Xlsx file to Html string with specific configurations.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml, ConverterConfig config)
        {
            ConvertXlsx(inputXlsx, outputHtml, config, null);
        }

        /// <summary>
        /// Convert a stream Xlsx file to Html string with progress callback event.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml, EventHandler<ConverterProgressCallbackEventArgs> progressCallback)
        {
            ConvertXlsx(inputXlsx, outputHtml, ConverterConfig.DefaultSettings, progressCallback);
        }

        /// <summary>
        /// Convert a stream Xlsx file to Html string with specific configurations and progress callback event.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml, ConverterConfig config, EventHandler<ConverterProgressCallbackEventArgs> progressCallback)
        {
            config = config ?? ConverterConfig.DefaultSettings;

            StreamWriter writer = new StreamWriter(outputHtml, config.Encoding, 65536);
            writer.BaseStream.Seek(0, SeekOrigin.Begin);
            writer.BaseStream.SetLength(0);

            try
            {
                writer.AutoFlush = true;

                if (!config.ConvertHtmlBodyOnly)
                {
                    writer.Write($@"<!DOCTYPE html>
<html>

<head>
    <meta charset=""UTF-8"">
    <title>{config.PageTitle}</title>

    <style>
        {config.PresetStyles}
    </style>
</head>
<body>");
                }
                else
                {
                    writer.Write($"<style>\n{config.PresetStyles}\n</style>");
                }

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(inputXlsx, false))
                {
                    WorkbookPart workbook = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbook.Workbook.Descendants<Sheet>();

                    WorkbookStylesPart styles = workbook.WorkbookStylesPart;
                    Stylesheet stylesheet = styles != null && styles.Stylesheet != null ? styles.Stylesheet : null;
                    Tuple<Dictionary<string, string>, string>[] htmlStyleCellFormats = new Tuple<Dictionary<string, string>, string>[stylesheet != null && stylesheet.CellFormats != null && stylesheet.CellFormats.HasChildren ? stylesheet.CellFormats.ChildElements.Count : 0];
                    for (int stylesheetFormatIndex = 0; stylesheetFormatIndex < htmlStyleCellFormats.Length; stylesheetFormatIndex++)
                    {
                        if (stylesheet.CellFormats.ChildElements[stylesheetFormatIndex] is CellFormat cellFormat)
                        {
                            Fill fill = (cellFormat.ApplyFill == null || (cellFormat.ApplyFill.HasValue && cellFormat.ApplyFill.Value)) && cellFormat.FillId != null && cellFormat.FillId.HasValue && stylesheet.Fills != null && stylesheet.Fills.HasChildren && cellFormat.FillId.Value < stylesheet.Fills.ChildElements.Count ? (Fill)stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value] : null;
                            Font font = (cellFormat.ApplyFont == null || (cellFormat.ApplyFont.HasValue && cellFormat.ApplyFont.Value)) && cellFormat.FontId != null && cellFormat.FontId.HasValue && stylesheet.Fonts != null && stylesheet.Fonts.HasChildren && cellFormat.FontId.Value < stylesheet.Fonts.ChildElements.Count ? (Font)stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value] : null;
                            Border border = (cellFormat.ApplyBorder == null || (cellFormat.ApplyBorder.HasValue && cellFormat.ApplyBorder.Value)) && cellFormat.BorderId != null && cellFormat.BorderId.HasValue && stylesheet.Borders != null && stylesheet.Borders.HasChildren && cellFormat.BorderId.Value < stylesheet.Borders.ChildElements.Count ? (Border)stylesheet.Borders.ChildElements[(int)cellFormat.BorderId.Value] : null;
                            htmlStyleCellFormats[stylesheetFormatIndex] = new Tuple<Dictionary<string, string>, string>(CellFormatToHtml(workbook, fill, font, border, cellFormat.ApplyAlignment == null || (cellFormat.ApplyAlignment.HasValue && cellFormat.ApplyAlignment.Value) ? cellFormat.Alignment : null, out string cellValueContainer), cellValueContainer);
                        }
                    }
                    Tuple<Dictionary<string, string>, string>[] htmlStyleDifferentialFormats = new Tuple<Dictionary<string, string>, string>[stylesheet != null && stylesheet.DifferentialFormats != null && stylesheet.DifferentialFormats.HasChildren ? stylesheet.DifferentialFormats.ChildElements.Count : 0];
                    for (int stylesheetFormatIndex = 0; stylesheetFormatIndex < htmlStyleDifferentialFormats.Length; stylesheetFormatIndex++)
                    {
                        if (stylesheet.DifferentialFormats.ChildElements[stylesheetFormatIndex] is DifferentialFormat differentialFormat)
                        {
                            htmlStyleDifferentialFormats[stylesheetFormatIndex] = new Tuple<Dictionary<string, string>, string>(CellFormatToHtml(workbook, differentialFormat.Fill, differentialFormat.Font, differentialFormat.Border, differentialFormat.Alignment, out string cellValueContainer), cellValueContainer);
                        }
                    }

                    IEnumerable<SharedStringTablePart> sharedStringTables = workbook.GetPartsOfType<SharedStringTablePart>();
                    SharedStringTable sharedStringTable = sharedStringTables.Any() ? sharedStringTables.First().SharedStringTable : null;
                    Tuple<string, string>[] cellValueSharedStrings = new Tuple<string, string>[sharedStringTable != null && sharedStringTable.HasChildren ? sharedStringTable.ChildElements.Count : 0];
                    for (int sharedStringIndex = 0; sharedStringIndex < cellValueSharedStrings.Length; sharedStringIndex++)
                    {
                        if (sharedStringTable.ChildElements[sharedStringIndex] is SharedStringItem sharedString)
                        {
                            string cellValue = string.Empty;
                            string cellValueRaw = string.Empty;
                            if (sharedString.HasChildren)
                            {
                                Run runLast = null;
                                foreach (OpenXmlElement element in sharedString.Descendants())
                                {
                                    if (element is Text text)
                                    {
                                        if (runLast == null || (runLast.Text != null && runLast.Text != text))
                                        {
                                            cellValue += GetEscapedString(text.Text);
                                            cellValueRaw += text.Text;
                                        }
                                        runLast = null;
                                    }
                                    else if (element is Run run && run.Text != null)
                                    {
                                        runLast = run;

                                        Dictionary<string, string> htmlStyleRun = new Dictionary<string, string>();
                                        if (config.ConvertStyle && run.RunProperties is RunProperties runProperties)
                                        {
                                            if (runProperties.GetFirstChild<RunFont>() is RunFont runFont && runFont.Val != null && runFont.Val.HasValue)
                                            {
                                                htmlStyleRun.Add("font-family", runFont.Val.Value);
                                            }
                                            htmlStyleRun = JoinHtmlAttributes(htmlStyleRun, FontToHtml(workbook, runProperties.GetFirstChild<Color>(), runProperties.GetFirstChild<FontSize>(), runProperties.GetFirstChild<Bold>(), runProperties.GetFirstChild<Italic>(), runProperties.GetFirstChild<Strike>(), runProperties.GetFirstChild<Underline>()));
                                        }

                                        string runContent = string.IsNullOrEmpty(run.Text.Text) ? run.Text.InnerText : run.Text.Text;
                                        cellValue += $"<p style=\"display: inline;{GetHtmlAttributesString(htmlStyleRun, true)}\">{GetEscapedString(runContent)}</p>";
                                        cellValueRaw += runContent;
                                    }
                                }
                            }
                            else
                            {
                                string text = sharedString.Text != null && !string.IsNullOrEmpty(sharedString.Text.Text) ? sharedString.Text.Text : sharedString.Text.InnerText;
                                cellValue = GetEscapedString(text);
                                cellValueRaw = text;
                            }

                            cellValueSharedStrings[sharedStringIndex] = new Tuple<string, string>(cellValue, cellValueRaw);
                        }
                    }

                    int sheetIndex = 0;
                    int sheetsCount = config.ConvertFirstSheetOnly ? Math.Min(sheets.Count(), 1) : sheets.Count();
                    foreach (Sheet currentSheet in sheets)
                    {
                        sheetIndex++;

                        if (config.ConvertFirstSheetOnly && sheetIndex > 1)
                        {
                            continue;
                        }
                        if (!config.ConvertHiddenSheet && currentSheet.State != null && currentSheet.State.HasValue && currentSheet.State.Value != SheetStateValues.Visible)
                        {
                            continue;
                        }

                        if (!(workbook.GetPartById(currentSheet.Id) is WorksheetPart worksheet))
                        {
                            continue;
                        }
                        Worksheet sheet = worksheet.Worksheet;

                        if (config.ConvertSheetNameTitle)
                        {
                            string tabColor = sheet.SheetProperties != null && sheet.SheetProperties.TabColor != null ? ColorTypeToHtml(workbook, sheet.SheetProperties.TabColor) : string.Empty;
                            writer.Write($"\n{new string(' ', 4)}<h5{(!string.IsNullOrEmpty(tabColor) ? " style=\"border-bottom-color: " + tabColor + ";\"" : string.Empty)}>{(currentSheet.Name != null && currentSheet.Name.HasValue ? currentSheet.Name.Value : "Untitled")}</h5>");
                        }

                        writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                        writer.Write($"\n{new string(' ', 8)}<table>");

                        double rowHeightDefault = sheet.SheetFormatProperties != null && sheet.SheetFormatProperties.DefaultRowHeight != null && sheet.SheetFormatProperties.DefaultRowHeight.HasValue ? sheet.SheetFormatProperties.DefaultRowHeight.Value / 0.75 : 20;

                        bool isMergeCellsContained = false;
                        List<MergeCellInfo> mergeCells = new List<MergeCellInfo>();
                        if (sheet.Descendants<MergeCells>().FirstOrDefault() is MergeCells mergeCellsGroup)
                        {
                            isMergeCellsContained = true;

                            foreach (MergeCell mergeCell in mergeCellsGroup.Cast<MergeCell>())
                            {
                                if (mergeCell.Reference == null || !mergeCell.Reference.HasValue)
                                {
                                    continue;
                                }

                                string[] range = mergeCell.Reference.Value.Split(':');
                                int firstColumn = GetColumnIndex(range[0]);
                                int secondColumn = GetColumnIndex(range[1]);
                                int firstRow = GetRowIndex(range[0]);
                                int secondRow = GetRowIndex(range[1]);

                                int fromColumn = Math.Min(firstColumn, secondColumn);
                                int toColumn = Math.Max(firstColumn, secondColumn);
                                int fromRow = Math.Min(firstRow, secondRow);
                                int toRow = Math.Max(firstRow, secondRow);

                                mergeCells.Add(new MergeCellInfo()
                                {
                                    FromColumn = fromColumn,
                                    FromRow = fromRow,
                                    ToColumn = toColumn,
                                    ToRow = toRow,
                                    ColumnSpanned = toColumn - fromColumn + 1,
                                    RowSpanned = toRow - fromRow + 1
                                });
                            }
                        }

                        ConditionalFormatting conditionalFormatting = null;
                        if (sheet.Elements<ConditionalFormatting>() is IEnumerable<ConditionalFormatting> conditionalFormattings && conditionalFormattings.Any())
                        {
                            conditionalFormatting = conditionalFormattings.First();
                        }

                        IEnumerable<Row> rows = sheet.Descendants<Row>();

                        int rowsCount = 0;
                        int columnsCount = 0;
                        if (sheet.SheetDimension != null && sheet.SheetDimension.Reference != null && sheet.SheetDimension.Reference.HasValue)
                        {
                            string[] dimension = sheet.SheetDimension.Reference.Value.Split(':');
                            int fromColumn = GetColumnIndex(dimension[0]);
                            int toColumn = GetColumnIndex(dimension[1]);
                            int fromRow = GetRowIndex(dimension[0]);
                            int toRow = GetRowIndex(dimension[1]);

                            rowsCount = toRow - fromColumn + 1;
                            columnsCount = toColumn - fromColumn + 1;
                        }
                        else
                        {
                            foreach (Cell cell in rows.SelectMany(x => x.Descendants<Cell>()))
                            {
                                if (cell.CellReference != null && cell.CellReference.HasValue)
                                {
                                    columnsCount = Math.Max(columnsCount, GetColumnIndex(cell.CellReference.Value) + 1);
                                    rowsCount = Math.Max(rowsCount, GetRowIndex(cell.CellReference.Value));

                                }
                            }
                        }

                        List<double> columnWidths = new List<double>(Enumerable.Repeat(double.NaN, columnsCount));
                        List<double> rowHeights = new List<double>(Enumerable.Repeat(rowHeightDefault, rowsCount));
                        if (sheet.GetFirstChild<Columns>() is Columns columnsGroup)
                        {
                            foreach (Column column in columnsGroup.Descendants<Column>())
                            {
                                for (int i = column.Min != null && column.Min.HasValue ? (int)column.Min.Value : 0; i <= (column.Max != null && column.Max.HasValue ? (int)column.Max.Value : 0); i++)
                                {
                                    if (column.CustomWidth != null && column.CustomWidth.HasValue && column.CustomWidth.Value && column.Width != null && column.Width.HasValue)
                                    {
                                        columnWidths[i - 1] = (column.Width.Value - 1) * 7 + 7;
                                    }
                                }
                            }
                        }

                        int rowIndex = 0;
                        int rowIndexLast = 0;
                        foreach (Row row in rows)
                        {
                            rowIndex++;

                            if (row.RowIndex == null || !row.RowIndex.HasValue)
                            {
                                continue;
                            }

                            if (row.RowIndex.Value - rowIndexLast > 1)
                            {
                                for (int i = 0; i < row.RowIndex.Value - rowIndexLast - 1; i++)
                                {
                                    writer.Write($"\n{new string(' ', 12)}<tr>");

                                    for (int j = 0; j < columnsCount; j++)
                                    {
                                        double actualCellWidth = j >= columnWidths.Count ? double.NaN : columnWidths[j];
                                        writer.Write($"\n{new string(' ', 16)}<td style=\"height: {rowHeightDefault}px; width: {(double.IsNaN(actualCellWidth) ? "auto" : actualCellWidth + "px")};\"></td>");
                                    }

                                    writer.Write($"\n{new string(' ', 12)}</tr>");
                                }
                            }

                            int columnIndex = 0;
                            double rowHeight = (row.CustomHeight == null || (row.CustomHeight.HasValue && row.CustomHeight.Value)) && row.Height != null && row.Height.HasValue ? row.Height.Value / 0.75 : rowHeightDefault;
                            rowHeights[(int)row.RowIndex.Value - 1] = rowHeight;

                            writer.Write($"\n{new string(' ', 12)}<tr>");

                            Cell[] cells = new Cell[columnsCount];
                            for (int i = 0; i < columnsCount; i++)
                            {
                                cells[i] = new Cell() { CellValue = new CellValue(string.Empty), CellReference = ((char)(64 + i + 1)).ToString() + row.RowIndex };
                            }
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                if (cell.CellReference != null && cell.CellReference.HasValue)
                                {
                                    cells[GetColumnIndex(cell.CellReference.Value)] = cell;
                                }
                            }

                            foreach (Cell cell in cells)
                            {
                                int columnSpanned = 1;
                                int rowSpanned = 1;

                                double cellHeightActual = config.ConvertSize ? rowHeight : double.NaN;
                                double cellWidthActual = config.ConvertSize ? (columnIndex >= columnWidths.Count ? double.NaN : columnWidths[columnIndex]) : double.NaN;

                                if (isMergeCellsContained && cell.CellReference != null && cell.CellReference.HasValue)
                                {
                                    int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
                                    int cellRowIndex = GetRowIndex(cell.CellReference.Value);

                                    if (mergeCells.Any(x => !(cellRowIndex == x.FromRow && cellColumnIndex == x.FromColumn) && cellRowIndex >= x.FromRow && cellRowIndex <= x.ToRow && cellColumnIndex >= x.FromColumn && cellColumnIndex <= x.ToColumn))
                                    {
                                        continue;
                                    }
                                    if (mergeCells.FirstOrDefault(x => x.FromColumn == cellColumnIndex && x.FromRow == cellRowIndex) is MergeCellInfo mergeCellInfo)
                                    {
                                        columnSpanned = mergeCellInfo.ColumnSpanned;
                                        rowSpanned = mergeCellInfo.RowSpanned;
                                        cellWidthActual = columnSpanned > 1 ? double.NaN : cellWidthActual;
                                        cellHeightActual = rowSpanned > 1 ? double.NaN : cellHeightActual;
                                    }
                                }

                                int styleIndex = cell.StyleIndex != null && cell.StyleIndex.HasValue ? (int)cell.StyleIndex.Value : (row.StyleIndex != null && row.StyleIndex.HasValue ? (int)row.StyleIndex.Value : -1);
                                CellFormat cellFormat = styleIndex >= 0 && stylesheet != null && stylesheet.CellFormats != null && stylesheet.CellFormats.HasChildren && styleIndex < stylesheet.CellFormats.ChildElements.Count && stylesheet.CellFormats.ChildElements[styleIndex] is CellFormat stylesheetCellFormat ? stylesheetCellFormat : null;

                                string numberFormatCode = string.Empty;
                                bool isNumberFormatDate = false;
                                switch (cellFormat != null && cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.HasValue && (cellFormat.ApplyNumberFormat == null || (cellFormat.ApplyNumberFormat.HasValue && cellFormat.ApplyNumberFormat.Value)) ? (int)cellFormat.NumberFormatId.Value : 0)
                                {
                                    case 0:
                                        numberFormatCode = string.Empty;
                                        break;
                                    case 1:
                                        numberFormatCode = "0";
                                        break;
                                    case 2:
                                        numberFormatCode = "0.00";
                                        break;
                                    case 3:
                                        numberFormatCode = "#,##0";
                                        break;
                                    case 4:
                                        numberFormatCode = "#,##0.00";
                                        break;
                                    case 9:
                                        numberFormatCode = "0%";
                                        break;
                                    case 10:
                                        numberFormatCode = "0.00%";
                                        break;
                                    case 11:
                                        numberFormatCode = "0.00E+00";
                                        break;
                                    case 12:
                                        numberFormatCode = "# ?/?";
                                        break;
                                    case 13:
                                        numberFormatCode = "# ??/??";
                                        break;
                                    case 14:
                                        numberFormatCode = "mm-dd-yy";
                                        isNumberFormatDate = true;
                                        break;
                                    case 15:
                                        numberFormatCode = "d-mmm-yy";
                                        isNumberFormatDate = true;
                                        break;
                                    case 16:
                                        numberFormatCode = "d-mmm";
                                        isNumberFormatDate = true;
                                        break;
                                    case 17:
                                        numberFormatCode = "mmm-yy";
                                        isNumberFormatDate = true;
                                        break;
                                    case 18:
                                        numberFormatCode = "h:mm AM/PM";
                                        isNumberFormatDate = true;
                                        break;
                                    case 19:
                                        numberFormatCode = "h:mm:ss AM/PM";
                                        isNumberFormatDate = true;
                                        break;
                                    case 20:
                                        numberFormatCode = "h:mm";
                                        isNumberFormatDate = true;
                                        break;
                                    case 21:
                                        numberFormatCode = "h:mm:ss";
                                        isNumberFormatDate = true;
                                        break;
                                    case 22:
                                        numberFormatCode = "m/d/yyh:mm";
                                        isNumberFormatDate = true;
                                        break;
                                    case 37:
                                        numberFormatCode = "#,##0 ;(#,##0)";
                                        break;
                                    case 38:
                                        numberFormatCode = "#,##0 ;[Red](#,##0)";
                                        break;
                                    case 39:
                                        numberFormatCode = "#,##0.00;(#,##0.00)";
                                        break;
                                    case 40:
                                        numberFormatCode = "#,##0.00;[Red](#,##0.00)";
                                        break;
                                    case 45:
                                        numberFormatCode = "mm:ss";
                                        break;
                                    case 46:
                                        numberFormatCode = "[h]:mm:ss";
                                        break;
                                    case 47:
                                        numberFormatCode = "mmss.0";
                                        break;
                                    case 48:
                                        numberFormatCode = "##0.0E+0";
                                        break;
                                    case 49:
                                        numberFormatCode = "@";
                                        break;
                                    default:
                                        if (stylesheet != null && stylesheet.NumberingFormats != null && stylesheet.NumberingFormats.HasChildren && stylesheet.NumberingFormats.ChildElements.FirstOrDefault(x => x is NumberingFormat item && item.NumberFormatId != null && item.NumberFormatId.HasValue && item.NumberFormatId.Value == cellFormat.NumberFormatId.Value) is NumberingFormat numberingFormat && numberingFormat.FormatCode != null && numberingFormat.FormatCode.HasValue)
                                        {
                                            numberFormatCode = numberingFormat.FormatCode.Value.Replace("&quot;", "\"");
                                        }
                                        break;
                                }

                                string cellValue = string.Empty;
                                string cellValueRaw = string.Empty;
                                if (cell.CellValue != null)
                                {
                                    cellValue = !string.IsNullOrEmpty(cell.CellValue.Text) ? cell.CellValue.Text : cell.CellValue.InnerText;

                                    bool isSharedString = false;
                                    if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.SharedString && int.TryParse(cellValue, out int sharedStringId) && sharedStringId >= 0 && sharedStringId < cellValueSharedStrings.Length && cellValueSharedStrings[sharedStringId] != null)
                                    {
                                        isSharedString = true;
                                        cellValue = cellValueSharedStrings[sharedStringId].Item1;
                                        cellValueRaw = cellValueSharedStrings[sharedStringId].Item2;
                                    }
                                    else
                                    {
                                        cellValueRaw = cellValue;
                                    }

                                    if (!string.IsNullOrEmpty(numberFormatCode))
                                    {
                                        if ((isNumberFormatDate || (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Date)) && double.TryParse(cellValueRaw, out double cellValueDate))
                                        {
                                            DateTime dateValue = DateTime.FromOADate(cellValueDate).Date;
                                            cellValue = GetEscapedString(dateValue.ToString(numberFormatCode.Replace("m", "M")));
                                        }
                                        else
                                        {
                                            string[] numberFormatCodeComponents = numberFormatCode.Split(';');
                                            if (numberFormatCodeComponents.Length > 1 && double.TryParse(cellValueRaw, out double cellValueNumber))
                                            {
                                                int indexComponent = cellValueNumber > 0 || (numberFormatCodeComponents.Length == 2 && cellValueNumber == 0) ? 0 : (cellValueNumber < 0 ? 1 : (numberFormatCodeComponents.Length >= 3 ? 2 : -1));
                                                numberFormatCode = indexComponent >= 0 ? numberFormatCodeComponents[indexComponent] : numberFormatCode;
                                            }
                                            else
                                            {
                                                numberFormatCode = numberFormatCodeComponents.Length >= 4 ? numberFormatCodeComponents[3] : numberFormatCode;
                                            }
                                            cellValue = GetEscapedString(GetFormattedNumber(cellValueRaw, numberFormatCode));
                                        }
                                    }
                                    else if (!isSharedString)
                                    {
                                        cellValue = GetEscapedString(cellValue);
                                    }
                                }

                                Dictionary<string, string> htmlStyleCell = new Dictionary<string, string>();
                                string cellValueContainer = "{0}";
                                if (config.ConvertStyle)
                                {
                                    if (!isNumberFormatDate)
                                    {
                                        string horizontalTextAlignment = GetGeneralAlignment(cell, cellValueRaw);
                                        if (horizontalTextAlignment != "left")
                                        {
                                            htmlStyleCell.Add("text-align", horizontalTextAlignment);
                                        }
                                    }
                                    if (styleIndex >= 0 && styleIndex < htmlStyleCellFormats.Length && htmlStyleCellFormats[styleIndex] != null)
                                    {
                                        htmlStyleCell = JoinHtmlAttributes(htmlStyleCell, htmlStyleCellFormats[styleIndex].Item1);
                                        cellValueContainer = cellValueContainer.Replace("{0}", htmlStyleCellFormats[styleIndex].Item2);
                                    }

                                    int differentialStyleIndex = -1;
                                    if (conditionalFormatting != null && false)
                                    {
                                        //TODO: check reference sequences
                                        int priorityCurrent = int.MaxValue;
                                        foreach (ConditionalFormattingRule formattingRule in conditionalFormatting.Elements<ConditionalFormattingRule>())
                                        {
                                            if (formattingRule == null || formattingRule.Priority == null || !formattingRule.Priority.HasValue || formattingRule.FormatId == null || !formattingRule.FormatId.HasValue || formattingRule.Priority.Value > priorityCurrent)
                                            {
                                                continue;
                                            }

                                            if (formattingRule.Type != null && formattingRule.Type.HasValue)
                                            {
                                                bool formattingCondition = false;
                                                if (formattingRule.Type.Value == ConditionalFormatValues.CellIs && formattingRule.GetFirstChild<Formula>() is Formula formula)
                                                {
                                                    string formulaText = formula.Text.Trim('"');
                                                    switch (formattingRule.Operator != null && formattingRule.Operator.HasValue ? formattingRule.Operator.Value : ConditionalFormattingOperatorValues.Equal)
                                                    {
                                                        case ConditionalFormattingOperatorValues.Equal:
                                                            formattingCondition = cellValueRaw == formulaText;
                                                            break;
                                                        case ConditionalFormattingOperatorValues.BeginsWith:
                                                            formattingCondition = cellValueRaw.StartsWith(formulaText);
                                                            break;
                                                        case ConditionalFormattingOperatorValues.EndsWith:
                                                            formattingCondition = cellValueRaw.EndsWith(formulaText);
                                                            break;
                                                        case ConditionalFormattingOperatorValues.ContainsText:
                                                            formattingCondition = cellValueRaw.Contains(formulaText);
                                                            break;
                                                    }
                                                }
                                                else if (formattingRule.Text != null && formattingRule.Text.HasValue)
                                                {
                                                    switch (formattingRule.Type.Value)
                                                    {
                                                        case ConditionalFormatValues.BeginsWith:
                                                            formattingCondition = cellValueRaw.StartsWith(formattingRule.Text.Value);
                                                            break;
                                                        case ConditionalFormatValues.EndsWith:
                                                            formattingCondition = cellValueRaw.EndsWith(formattingRule.Text.Value);
                                                            break;
                                                        case ConditionalFormatValues.ContainsText:
                                                            formattingCondition = cellValueRaw.Contains(formattingRule.Text.Value);
                                                            break;
                                                        case ConditionalFormatValues.NotContainsText:
                                                            formattingCondition = !cellValueRaw.Contains(formattingRule.Text.Value);
                                                            break;
                                                    }
                                                }

                                                if (formattingCondition)
                                                {
                                                    differentialStyleIndex = (int)formattingRule.FormatId.Value;
                                                }
                                            }

                                            priorityCurrent = formattingRule.Priority.Value;
                                        }
                                    }
                                    if (differentialStyleIndex >= 0 && differentialStyleIndex < htmlStyleDifferentialFormats.Length && htmlStyleDifferentialFormats[differentialStyleIndex] != null)
                                    {
                                        htmlStyleCell = JoinHtmlAttributes(htmlStyleCell, htmlStyleDifferentialFormats[differentialStyleIndex].Item1);
                                        cellValueContainer = cellValueContainer.Replace("{0}", htmlStyleDifferentialFormats[differentialStyleIndex].Item2);
                                    }
                                }

                                writer.Write($"\n{new string(' ', 16)}<td colspan=\"{columnSpanned}\" rowspan=\"{rowSpanned}\" style=\"height: {(double.IsNaN(cellHeightActual) ? "auto" : cellHeightActual + "px")}; width: {(double.IsNaN(cellWidthActual) ? "auto" : cellWidthActual + "px")};{GetHtmlAttributesString(htmlStyleCell, true)}\">{cellValueContainer.Replace("{0}", cellValue)}</td>");

                                columnIndex += columnSpanned;
                            }

                            writer.Write($"\n{new string(' ', 12)}</tr>");

                            rowIndexLast = (int)row.RowIndex.Value;

                            progressCallback?.Invoke(document, new ConverterProgressCallbackEventArgs(sheetIndex, sheetsCount, rowIndex, rowsCount));
                        }

                        writer.Write($"\n{new string(' ', 8)}</table>");

                        if (worksheet.DrawingsPart != null && worksheet.DrawingsPart.WorksheetDrawing != null)
                        {
                            //TODO: position scaled-adjustments
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absoluteAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor>())
                            {
                                double left = absoluteAnchor.Position != null && absoluteAnchor.Position.X != null && absoluteAnchor.Position.X.HasValue ? (double)absoluteAnchor.Position.X.Value / 2 * 96 / 72 : double.NaN;
                                double top = absoluteAnchor.Position != null && absoluteAnchor.Position.Y != null && absoluteAnchor.Position.Y.HasValue ? (double)absoluteAnchor.Position.Y.Value / 2 * 96 / 72 : double.NaN;
                                double width = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cx != null && absoluteAnchor.Extent.Cx.HasValue ? (double)absoluteAnchor.Extent.Cx.Value / 914400 * 96 : double.NaN;
                                double height = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cy != null && absoluteAnchor.Extent.Cy.HasValue ? (double)absoluteAnchor.Extent.Cy.Value / 914400 * 96 : double.NaN;
                                DrawingsToHtml(worksheet, absoluteAnchor, writer, left, top, width, height, config.ConvertPicture);
                            }
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>())
                            {
                                double left = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) ? columnWidths.Take(Math.Min(columnWidths.Count, columnId)).Sum() + (oneCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(oneCellAnchor.FromMarker.ColumnOffset.Text, out double columnOffset) ? columnOffset / 914400 * 96 : 0) : double.NaN;
                                double top = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) ? rowHeights.Take(Math.Min(rowHeights.Count, rowId)).Sum() + (oneCellAnchor.FromMarker.RowOffset != null && double.TryParse(oneCellAnchor.FromMarker.RowOffset.Text, out double rowOffset) ? rowOffset / 914400 * 96 : 0) : double.NaN;
                                double width = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cx != null && oneCellAnchor.Extent.Cx.HasValue ? (double)oneCellAnchor.Extent.Cx.Value / 914400 * 96 : double.NaN;
                                double height = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cy != null && oneCellAnchor.Extent.Cy.HasValue ? (double)oneCellAnchor.Extent.Cy.Value / 914400 * 96 : double.NaN;
                                DrawingsToHtml(worksheet, oneCellAnchor, writer, left, top, width, height, config.ConvertPicture);
                            }
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                            {
                                double fromLeft = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) ? columnWidths.Take(Math.Min(columnWidths.Count, fromColumnId)).Sum() + (twoCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.FromMarker.ColumnOffset.Text, out double fromColumnOffset) ? fromColumnOffset / 914400 * 96 : 0) : double.NaN;
                                double fromTop = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) ? rowHeights.Take(Math.Min(rowHeights.Count, fromRowId)).Sum() + (twoCellAnchor.FromMarker.RowOffset != null && double.TryParse(twoCellAnchor.FromMarker.RowOffset.Text, out double fromRowOffset) ? fromRowOffset / 914400 * 96 : 0) : double.NaN;
                                double toLeft = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) ? columnWidths.Take(Math.Min(columnWidths.Count, toColumnId)).Sum() + (twoCellAnchor.ToMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.ToMarker.ColumnOffset.Text, out double toColumnOffset) ? toColumnOffset / 914400 * 96 : 0) : double.NaN;
                                double toTop = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) ? rowHeights.Take(Math.Min(rowHeights.Count, toRowId)).Sum() + (twoCellAnchor.ToMarker.RowOffset != null && double.TryParse(twoCellAnchor.ToMarker.RowOffset.Text, out double toRowOffset) ? toRowOffset / 914400 * 96 : 0) : double.NaN;
                                DrawingsToHtml(worksheet, twoCellAnchor, writer, Math.Min(fromLeft, toLeft), Math.Min(fromTop, toTop), Math.Abs(toLeft - fromLeft), Math.Abs(toTop - fromTop), config.ConvertPicture);
                            }
                        }

                        writer.Write($"\n{new string(' ', 4)}</div>");
                    }
                }

                if (!config.ConvertHtmlBodyOnly)
                {
                    writer.Write("\n</body>\n</html>");
                }
            }
            catch (Exception ex)
            {
                writer.BaseStream.Seek(0, SeekOrigin.Begin);
                writer.BaseStream.SetLength(0);
                writer.Write(config.ErrorMessage.Replace("{EXCEPTION}", ex.Message));
            }
            finally
            {
                writer.BaseStream.Seek(0, SeekOrigin.Begin);
                writer.Close();
                writer.Dispose();
            }
        }

        #endregion

        #region Private Methods

        private static int GetColumnIndex(string cellName)
        {
            int columnNumber = -1;
            Match match = regexLetters.Match(cellName);
            if (match.Success)
            {
                int mulitplier = 1;
                foreach (char c in match.Value.ToUpper().ToCharArray().Reverse())
                {
                    columnNumber += mulitplier * (c - 64);
                    mulitplier *= 26;
                }
            }
            return Math.Max(0, columnNumber);
        }

        private static int GetRowIndex(string cellName)
        {
            Match match = regexNumbers.Match(cellName);
            return match.Success && int.TryParse(match.Value, out int rowIndex) ? rowIndex : 0;
        }

        private static string GetEscapedString(string value)
        {
            return System.Web.HttpUtility.HtmlEncode(value).Replace(" ", "&nbsp;");
        }

        private static string GetHtmlAttributesString(Dictionary<string, string> attributes, bool isAdditional)
        {
            string htmlAttributes = string.Empty;
            foreach (KeyValuePair<string, string> pair in attributes)
            {
                if (!string.IsNullOrEmpty(pair.Key) && !string.IsNullOrEmpty(pair.Value))
                {
                    htmlAttributes += $"{pair.Key}: {pair.Value}; ";
                }
            }
            return isAdditional ? htmlAttributes : htmlAttributes.Trim();
        }

        private static Dictionary<string, string> JoinHtmlAttributes(Dictionary<string, string> original, Dictionary<string, string> joined)
        {
            if (joined == null)
            {
                return original;
            }

            foreach (KeyValuePair<string, string> pair in joined)
            {
                if (original.ContainsKey(pair.Key))
                {
                    original[pair.Key] = pair.Value;
                }
                else
                {
                    original.Add(pair.Key, pair.Value);
                }
            }
            return original;
        }

        private static string GetFormattedNumber(string value, string format)
        {
            if (string.IsNullOrEmpty(format))
            {
                return value;
            }

            bool isValueNumber = double.TryParse(value, out double valueNumber);
            if (!isValueNumber && !format.Contains("@"))
            {
                return value;
            }

            int[] indexes = new int[6] { value.Length, format.Length, format.Length, format.Length, 0, 0 };
            bool isPeriodRequired = false;
            Action actionUpdateValue = () =>
            {
                value = isValueNumber ? valueNumber.ToString() : value;
                indexes[0] = value.IndexOf('.');
                indexes[0] = indexes[0] < 0 ? value.Length : indexes[0];
            };
            actionUpdateValue.Invoke();

            string scientificPower = "0";
            bool isScientificNegative = false;
            bool isFormattingScientific = false;

            bool isIncreasing = true;
            bool isFormatting = false;
            string result = string.Empty;
            string resultFormatted = string.Empty;
            while (indexes[5] < format.Length || !isFormatting)
            {
                if (indexes[5] >= format.Length)
                {
                    indexes[2] = Math.Min(indexes[3] + 1, indexes[2]);

                    isIncreasing = false;
                    indexes[4] = indexes[0];
                    indexes[5] = indexes[2] - 1;
                    isFormatting = true;
                    continue;
                }
                else if (indexes[5] < 0)
                {
                    result = new string(result.Reverse().ToArray());
                    isIncreasing = true;
                    indexes[4] = indexes[0];
                    indexes[5] = indexes[2];
                    continue;
                }

                char formatChar = format[indexes[5]];
                if ((isIncreasing && indexes[5] + 1 < format.Length && formatChar == '\\') || (!isIncreasing && indexes[5] > 0 && format[indexes[5] - 1] == '\\'))
                {
                    result += isFormatting ? format[isIncreasing ? indexes[5] + 1 : indexes[5]].ToString() : string.Empty;
                    indexes[5] += isIncreasing ? 2 : -2;
                    continue;
                }
                else if (isIncreasing ? formatChar == '[' && indexes[5] + 1 < format.Length : formatChar == ']' && indexes[5] > 0)
                {
                    do
                    {
                        //TODO: conditions
                        indexes[5] += isIncreasing ? 1 : -1;
                    } while (isIncreasing ? indexes[5] + 1 < format.Length && format[indexes[5] + 1] != ']' : indexes[5] > 0 && format[indexes[5] - 1] != '[');
                    indexes[5] += isIncreasing ? 2 : -2;
                    continue;
                }
                else if (formatChar == '\"' && (isIncreasing ? indexes[5] + 1 < format.Length : indexes[5] > 0))
                {
                    do
                    {
                        indexes[5] += isIncreasing ? 1 : -1;
                        result += isFormatting ? format[indexes[5]].ToString() : string.Empty;
                    }
                    while (isIncreasing ? indexes[5] + 1 < format.Length && format[indexes[5] + 1] != '\"' : indexes[5] > 0 && format[indexes[5] - 1] != '\"');
                    indexes[5] += isIncreasing ? 2 : -2;
                    continue;
                }
                else if ((isIncreasing && indexes[5] + 1 < format.Length && formatChar == '*') || (!isIncreasing && indexes[5] > 0 && format[indexes[5] - 1] == '*'))
                {
                    result += isFormatting ? format[isIncreasing ? indexes[5] + 1 : indexes[5]].ToString() : string.Empty;
                    indexes[5] += isIncreasing ? 2 : -2;
                    continue;
                }
                else if ((isIncreasing && indexes[5] + 1 < format.Length && formatChar == '_') || (!isIncreasing && indexes[5] > 0 && format[indexes[5] - 1] == '_'))
                {
                    result += isFormatting ? " " : string.Empty;
                    indexes[5] += isIncreasing ? 2 : -2;
                    continue;
                }
                else if (isFormatting && !isIncreasing && indexes[5] > 0 && format[indexes[5] - 1] == 'E' && (formatChar == '+' || formatChar == '-'))
                {
                    result = resultFormatted + new string(result.Reverse().ToArray());
                    isIncreasing = true;
                    indexes[5] = indexes[3] + 1;
                    isFormattingScientific = false;
                    continue;
                }

                if (!isFormatting && isValueNumber)
                {
                    if (formatChar == '.')
                    {
                        indexes[2] = Math.Min(indexes[5], indexes[2]);
                    }
                    else if (formatChar == '0' || formatChar == '#' || formatChar == '?')
                    {
                        indexes[1] = Math.Min(indexes[5], indexes[1]);
                        indexes[3] = indexes[5];
                        isPeriodRequired = (indexes[5] > indexes[2] && (formatChar == '0' || formatChar == '?') && scientificPower == "0") || isPeriodRequired;
                    }
                    else if (formatChar == '%')
                    {
                        valueNumber *= 100;
                        actionUpdateValue.Invoke();
                    }
                    else if (formatChar == 'E' && isIncreasing && indexes[5] + 1 < format.Length && (format[indexes[5] + 1] == '+' || format[indexes[5] + 1] == '-'))
                    {
                        if (indexes[0] > 1)
                        {
                            isScientificNegative = false;
                            scientificPower = (indexes[0] - 1).ToString();
                            valueNumber /= Math.Pow(10, indexes[0] - 1);
                            actionUpdateValue.Invoke();
                        }
                        else if (indexes[0] > 0 && value.Length > indexes[0] && value[0] == '0')
                        {
                            int digit = 0;
                            for (int i = indexes[0] + 1; i < value.Length; i++)
                            {
                                if (value[i] != '0')
                                {
                                    digit = i;
                                    break;
                                }
                            }
                            if (digit > indexes[0])
                            {
                                isScientificNegative = true;
                                scientificPower = (digit - indexes[0]).ToString();
                                valueNumber *= Math.Pow(10, digit - indexes[0]);
                                actionUpdateValue.Invoke();
                            }
                        }
                        indexes[2] = Math.Min(indexes[5], indexes[2]);
                        indexes[5]++;
                    }
                    else if (formatChar == '/')
                    {
                        //TODO: fractions
                        double valueAbsolute = Math.Abs(valueNumber);
                        int valueFloor = (int)Math.Floor(valueAbsolute);
                        valueAbsolute -= valueFloor;

                        int fractionNumerator = 1;
                        int fractionDenominator = 1;
                        double maxError = valueAbsolute * 0.001;
                        if (valueAbsolute == 0)
                        {
                            fractionNumerator = 0;
                            fractionDenominator = 1;
                        }
                        else if (valueAbsolute < maxError)
                        {
                            fractionNumerator = valueNumber > 0 ? valueFloor : -valueFloor;
                            fractionDenominator = 1;
                        }
                        else if (1 - maxError < valueAbsolute)
                        {
                            fractionNumerator = valueNumber > 0 ? (valueFloor + 1) : -(valueFloor + 1);
                            fractionDenominator = 1;
                        }
                        else
                        {
                            int[] fractionParts = new int[4] { 0, 1, 1, 1 };
                            Func<int, int, int, int, Func<int, int, bool>, bool> actionFindNewValue = (indexNumerator, indexDenominator, incrementNumerator, incrementDenominator, function) =>
                            {
                                fractionParts[indexNumerator] += incrementNumerator;
                                fractionParts[indexDenominator] += incrementDenominator;
                                if (function.Invoke(fractionParts[indexNumerator], fractionParts[indexDenominator]))
                                {
                                    int weight = 1;
                                    do
                                    {
                                        weight *= 2;
                                        fractionParts[indexNumerator] += incrementNumerator * weight;
                                        fractionParts[indexDenominator] += incrementDenominator * weight;
                                    }
                                    while (function.Invoke(fractionParts[indexNumerator], fractionParts[indexDenominator]));
                                    do
                                    {
                                        weight /= 2;
                                        int decrementNumerator = incrementNumerator * weight;
                                        int decrementDenominator = incrementDenominator * weight;
                                        if (!function.Invoke(fractionParts[indexNumerator] - decrementNumerator, fractionParts[indexDenominator] - decrementDenominator))
                                        {
                                            fractionParts[indexNumerator] -= decrementNumerator;
                                            fractionParts[indexDenominator] -= decrementDenominator;
                                        }
                                    }
                                    while (weight > 1);
                                }
                                return true;
                            };

                            while (true)
                            {
                                int middleNumerator = fractionParts[0] + fractionParts[2];
                                int middleDenominator = fractionParts[1] + fractionParts[3];
                                if (middleDenominator * (valueAbsolute + maxError) < middleNumerator)
                                {
                                    actionFindNewValue(2, 3, fractionParts[0], fractionParts[1], (numerator, denominator) => (fractionParts[1] + denominator) * (valueAbsolute + maxError) < (fractionParts[0] + numerator));
                                }
                                else if (middleNumerator < (valueAbsolute - maxError) * middleDenominator)
                                {
                                    actionFindNewValue(0, 1, fractionParts[2], fractionParts[3], (numerator, denominator) => (numerator + fractionParts[2]) < (valueAbsolute - maxError) * (denominator + fractionParts[3]));
                                }
                                else
                                {
                                    fractionNumerator = valueNumber > 0 ? valueFloor * middleDenominator + middleNumerator : -(valueFloor * middleDenominator + middleNumerator);
                                    fractionDenominator = middleDenominator;
                                    break;
                                }
                            }
                        }

                        bool debug = true;
                    }
                }
                else if (isFormatting)
                {
                    if (formatChar == '@')
                    {
                        result += isIncreasing ? value : new string(value.Reverse().ToArray());
                    }
                    else if (isValueNumber && formatChar == '.')
                    {
                        if (isPeriodRequired || (isIncreasing && indexes[4] + 1 < value.Length))
                        {
                            result += ".";
                        }
                    }
                    else if (isValueNumber && formatChar == ',')
                    {
                        if (isIncreasing ? indexes[4] + 1 < value.Length : indexes[4] > 0)
                        {
                            result += ",";
                        }
                    }
                    else if (isValueNumber && formatChar == 'E' && isIncreasing && indexes[5] + 1 < format.Length && (format[indexes[5] + 1] == '+' || format[indexes[5] + 1] == '-'))
                    {
                        resultFormatted = result + (!isScientificNegative ? (format[indexes[5] + 1] == '-' ? "E" : "E+") : "E-");
                        result = string.Empty;
                        isIncreasing = false;
                        indexes[4] = scientificPower.Length;
                        indexes[5] = indexes[3];
                        isFormattingScientific = true;
                        continue;
                    }
                    else if (isValueNumber && (formatChar == '0' || formatChar == '#' || formatChar == '?'))
                    {
                        indexes[4] += isIncreasing ? 1 : -1;
                        if (indexes[4] >= 0 && indexes[4] < (!isFormattingScientific ? value.Length : scientificPower.Length) && (formatChar == '0' || indexes[4] > 0 || value[indexes[4]] != '0' || isPeriodRequired || isFormattingScientific))
                        {
                            if (isIncreasing && (indexes[5] >= indexes[3] || (indexes[5] + 2 < format.Length && format[indexes[5] + 1] == 'E' && (format[indexes[5] + 2] == '+' || format[indexes[5] + 2] == '-'))) && indexes[4] + 1 < value.Length && int.TryParse(value[indexes[4] + 1].ToString(), out int next) && next >= 5)
                            {
                                return GetFormattedNumber((valueNumber + (11 - next) / Math.Pow(10, indexes[4] + 1 - indexes[0])).ToString(), format);
                            }

                            result += !isFormattingScientific ? value[indexes[4]].ToString() : scientificPower[indexes[4]].ToString();
                            if (!isFormattingScientific ? indexes[5] <= indexes[1] : (indexes[5] - 2 >= 0 && format[indexes[5] - 2] == 'E' && (format[indexes[5] - 1] == '+' || format[indexes[5] - 1] == '-')))
                            {
                                result += new string((!isFormattingScientific ? value : scientificPower).Substring(0, indexes[4]).Reverse().ToArray());
                            }
                        }
                        else
                        {
                            result += formatChar == '0' ? "0" : (formatChar == '?' ? " " : string.Empty);
                        }
                    }
                    else
                    {
                        result += formatChar;
                    }
                }

                indexes[5] += isIncreasing ? 1 : -1;
            }

            return result;
        }

        private static Dictionary<string, string> CellFormatToHtml(WorkbookPart workbook, Fill fill, Font font, Border border, Alignment alignment, out string cellValueContainer)
        {
            Dictionary<string, string> htmlStyle = new Dictionary<string, string>();
            cellValueContainer = "{0}";

            if (fill != null && fill.PatternFill != null && fill.PatternFill.PatternType != null && fill.PatternFill.PatternType.HasValue && fill.PatternFill.PatternType.Value != PatternValues.None)
            {
                string background = string.Empty;
                if (fill.PatternFill.ForegroundColor != null)
                {
                    background = ColorTypeToHtml(workbook, fill.PatternFill.ForegroundColor);
                }
                if (string.IsNullOrEmpty(background) && fill.PatternFill.BackgroundColor != null)
                {
                    background = ColorTypeToHtml(workbook, fill.PatternFill.BackgroundColor);
                }
                if (!string.IsNullOrEmpty(background))
                {
                    htmlStyle.Add("background-color", background);
                }
            }

            if (font != null)
            {
                if (font.FontName != null && font.FontName.Val != null && font.FontName.Val.HasValue)
                {
                    htmlStyle.Add("font-family", font.FontName.Val.Value);
                }
                htmlStyle = JoinHtmlAttributes(htmlStyle, FontToHtml(workbook, font.Color, font.FontSize, font.Bold, font.Italic, font.Strike, font.Underline));
            }

            if (border != null)
            {
                string[] borderWidths = new string[4] { "revert", "revert", "revert", "revert" };
                string[] borderStyles = new string[4] { "revert", "revert", "revert", "revert" };
                string[] borderColors = new string[4] { "revert", "revert", "revert", "revert" };
                if (border.LeftBorder != null)
                {
                    BorderPropertiesToHtml(workbook, border.LeftBorder, ref borderWidths[0], ref borderStyles[0], ref borderColors[0]);
                }
                if (border.RightBorder != null)
                {
                    BorderPropertiesToHtml(workbook, border.RightBorder, ref borderWidths[1], ref borderStyles[1], ref borderColors[1]);
                }
                if (border.TopBorder != null)
                {
                    BorderPropertiesToHtml(workbook, border.TopBorder, ref borderWidths[2], ref borderStyles[2], ref borderColors[2]);
                }
                if (border.BottomBorder != null)
                {
                    BorderPropertiesToHtml(workbook, border.BottomBorder, ref borderWidths[3], ref borderStyles[3], ref borderColors[3]);
                }

                if (borderWidths.Any(x => x != "revert"))
                {
                    htmlStyle.Add("border-width", $"{borderWidths[0]} {borderWidths[1]} {borderWidths[2]} {borderWidths[3]}");
                }
                if (borderStyles.Any(x => x != "revert"))
                {
                    htmlStyle.Add("border-style", $"{borderStyles[0]} {borderStyles[1]} {borderStyles[2]} {borderStyles[3]}");
                }
                if (borderColors.Any(x => x != "revert"))
                {
                    htmlStyle.Add("border-color", $"{borderColors[0]} {borderColors[1]} {borderColors[2]} {borderColors[3]}");
                }
            }

            if (alignment != null)
            {
                string verticalTextAlignment = "bottom";
                string horizontalTextAlignment = "left";

                if (alignment.Vertical != null && alignment.Vertical.HasValue)
                {
                    switch (alignment.Vertical.Value)
                    {
                        case VerticalAlignmentValues.Bottom:
                            verticalTextAlignment = "bottom";
                            break;
                        case VerticalAlignmentValues.Center:
                            verticalTextAlignment = "middle";
                            break;
                        case VerticalAlignmentValues.Top:
                            verticalTextAlignment = "top";
                            break;
                    }
                }
                if (alignment.Horizontal != null && alignment.Horizontal.HasValue)
                {
                    switch (alignment.Horizontal.Value)
                    {
                        case HorizontalAlignmentValues.Left:
                            horizontalTextAlignment = "left";
                            break;
                        case HorizontalAlignmentValues.Center:
                            horizontalTextAlignment = "center";
                            break;
                        case HorizontalAlignmentValues.Right:
                            horizontalTextAlignment = "right";
                            break;
                        case HorizontalAlignmentValues.Justify:
                            horizontalTextAlignment = "justify";
                            break;
                    }
                }

                htmlStyle.Add("text-align", horizontalTextAlignment);
                htmlStyle.Add("vertical-align", verticalTextAlignment);

                if (alignment.WrapText != null && alignment.WrapText.HasValue && alignment.WrapText.Value)
                {
                    htmlStyle.Add("word-wrap", "break-word");
                    htmlStyle.Add("white-space", "normal");
                }
                if (alignment.TextRotation != null && alignment.TextRotation.HasValue)
                {
                    cellValueContainer = $"<div style=\"width: fit-content; rotate: -{alignment.TextRotation.Value}deg;\">{{0}}</div>";
                }
            }

            return htmlStyle;
        }

        private static string ColorTypeToHtml(WorkbookPart workbook, ColorType type)
        {
            if (type == null)
            {
                return string.Empty;
            }

            RgbaColor rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };

            if (type.Rgb != null && type.Rgb.HasValue)
            {
                rgbColor = HexToRgba(type.Rgb.Value);
            }
            else if (type.Indexed != null && type.Indexed.HasValue)
            {
                switch (type.Indexed.Value)
                {
                    case 0:
                        rgbColor = HexToRgba("#000000");
                        break;
                    case 1:
                        rgbColor = HexToRgba("#FFFFFF");
                        break;
                    case 2:
                        rgbColor = HexToRgba("#FF0000");
                        break;
                    case 3:
                        rgbColor = HexToRgba("#00FF00");
                        break;
                    case 4:
                        rgbColor = HexToRgba("#0000FF");
                        break;
                    case 5:
                        rgbColor = HexToRgba("#FFFF00");
                        break;
                    case 6:
                        rgbColor = HexToRgba("#FF00FF");
                        break;
                    case 7:
                        rgbColor = HexToRgba("#00FFFF");
                        break;
                    case 8:
                        rgbColor = HexToRgba("#000000");
                        break;
                    case 9:
                        rgbColor = HexToRgba("#FFFFFF");
                        break;
                    case 10:
                        rgbColor = HexToRgba("#FF0000");
                        break;
                    case 11:
                        rgbColor = HexToRgba("#00FF00");
                        break;
                    case 12:
                        rgbColor = HexToRgba("#0000FF");
                        break;
                    case 13:
                        rgbColor = HexToRgba("#FFFF00");
                        break;
                    case 14:
                        rgbColor = HexToRgba("#FF00FF");
                        break;
                    case 15:
                        rgbColor = HexToRgba("#00FFFF");
                        break;
                    case 16:
                        rgbColor = HexToRgba("#800000");
                        break;
                    case 17:
                        rgbColor = HexToRgba("#008000");
                        break;
                    case 18:
                        rgbColor = HexToRgba("#000080");
                        break;
                    case 19:
                        rgbColor = HexToRgba("#808000");
                        break;
                    case 20:
                        rgbColor = HexToRgba("#800080");
                        break;
                    case 21:
                        rgbColor = HexToRgba("#008080");
                        break;
                    case 22:
                        rgbColor = HexToRgba("#C0C0C0");
                        break;
                    case 23:
                        rgbColor = HexToRgba("#808080");
                        break;
                    case 24:
                        rgbColor = HexToRgba("#9999FF");
                        break;
                    case 25:
                        rgbColor = HexToRgba("#993366");
                        break;
                    case 26:
                        rgbColor = HexToRgba("#FFFFCC");
                        break;
                    case 27:
                        rgbColor = HexToRgba("#CCFFFF");
                        break;
                    case 28:
                        rgbColor = HexToRgba("#660066");
                        break;
                    case 29:
                        rgbColor = HexToRgba("#FF8080");
                        break;
                    case 30:
                        rgbColor = HexToRgba("#0066CC");
                        break;
                    case 31:
                        rgbColor = HexToRgba("#CCCCFF");
                        break;
                    case 32:
                        rgbColor = HexToRgba("#000080");
                        break;
                    case 33:
                        rgbColor = HexToRgba("#FF00FF");
                        break;
                    case 34:
                        rgbColor = HexToRgba("#FFFF00");
                        break;
                    case 35:
                        rgbColor = HexToRgba("#00FFFF");
                        break;
                    case 36:
                        rgbColor = HexToRgba("#800080");
                        break;
                    case 37:
                        rgbColor = HexToRgba("#800000");
                        break;
                    case 38:
                        rgbColor = HexToRgba("#008080");
                        break;
                    case 39:
                        rgbColor = HexToRgba("#0000FF");
                        break;
                    case 40:
                        rgbColor = HexToRgba("#00CCFF");
                        break;
                    case 41:
                        rgbColor = HexToRgba("#CCFFFF");
                        break;
                    case 42:
                        rgbColor = HexToRgba("#CCFFCC");
                        break;
                    case 43:
                        rgbColor = HexToRgba("#FFFF99");
                        break;
                    case 44:
                        rgbColor = HexToRgba("#99CCFF");
                        break;
                    case 45:
                        rgbColor = HexToRgba("#FF99CC");
                        break;
                    case 46:
                        rgbColor = HexToRgba("#CC99FF");
                        break;
                    case 47:
                        rgbColor = HexToRgba("#FFCC99");
                        break;
                    case 48:
                        rgbColor = HexToRgba("#3366FF");
                        break;
                    case 49:
                        rgbColor = HexToRgba("#33CCCC");
                        break;
                    case 50:
                        rgbColor = HexToRgba("#99CC00");
                        break;
                    case 51:
                        rgbColor = HexToRgba("#FFCC00");
                        break;
                    case 52:
                        rgbColor = HexToRgba("#FF9900");
                        break;
                    case 53:
                        rgbColor = HexToRgba("#FF6600");
                        break;
                    case 54:
                        rgbColor = HexToRgba("#666699");
                        break;
                    case 55:
                        rgbColor = HexToRgba("#969696");
                        break;
                    case 56:
                        rgbColor = HexToRgba("#003366");
                        break;
                    case 57:
                        rgbColor = HexToRgba("#339966");
                        break;
                    case 58:
                        rgbColor = HexToRgba("#003300");
                        break;
                    case 59:
                        rgbColor = HexToRgba("#333300");
                        break;
                    case 60:
                        rgbColor = HexToRgba("#993300");
                        break;
                    case 61:
                        rgbColor = HexToRgba("#993366");
                        break;
                    case 62:
                        rgbColor = HexToRgba("#333399");
                        break;
                    case 63:
                        rgbColor = HexToRgba("#333333");
                        break;
                    case 64:
                        rgbColor = HexToRgba("#808080");
                        break;
                    case 65:
                        rgbColor = HexToRgba("#FFFFFF");
                        break;
                    default:
                        return string.Empty;
                }
            }
            else if (type.Theme != null && type.Theme.HasValue && workbook.ThemePart != null && workbook.ThemePart.Theme != null && workbook.ThemePart.Theme.ThemeElements != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme.HasChildren && type.Theme.Value < workbook.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements.Count && workbook.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)type.Theme.Value] is DocumentFormat.OpenXml.Drawing.Color2Type color)
            {
                //TODO: correct colors
                if (color.RgbColorModelHex != null && color.RgbColorModelHex.Val != null && color.RgbColorModelHex.Val.HasValue)
                {
                    rgbColor = HexToRgba(color.RgbColorModelHex.Val.Value);
                }
                else if (color.RgbColorModelPercentage != null)
                {
                    rgbColor.R = color.RgbColorModelPercentage.RedPortion.HasValue ? (int)(color.RgbColorModelPercentage.RedPortion.Value / 100000.0 * 255) : 0;
                    rgbColor.G = color.RgbColorModelPercentage.GreenPortion.HasValue ? (int)(color.RgbColorModelPercentage.GreenPortion.Value / 100000.0 * 255) : 0;
                    rgbColor.B = color.RgbColorModelPercentage.BluePortion.HasValue ? (int)(color.RgbColorModelPercentage.BluePortion.Value / 100000.0 * 255) : 0;
                }
                else if (color.HslColor != null)
                {
                    HlsToRgb(color.HslColor.HueValue.HasValue ? color.HslColor.HueValue.Value : 0, color.HslColor.LumValue.HasValue ? color.HslColor.LumValue.Value : 0, color.HslColor.SatValue.HasValue ? color.HslColor.SatValue.Value : 0, out int r, out int g, out int b);
                    rgbColor.R = r;
                    rgbColor.G = g;
                    rgbColor.B = b;
                }
                else if (color.SystemColor != null && color.SystemColor.Val != null && color.SystemColor.Val.HasValue)
                {
                    switch (color.SystemColor.Val.Value)
                    {
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveBorder:
                            rgbColor = new RgbaColor() { R = 180, G = 180, B = 180, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveCaption:
                            rgbColor = new RgbaColor() { R = 153, G = 180, B = 209, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ApplicationWorkspace:
                            rgbColor = new RgbaColor() { R = 171, G = 171, B = 171, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.Background:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonFace:
                            rgbColor = new RgbaColor() { R = 240, G = 240, B = 240, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonHighlight:
                            rgbColor = new RgbaColor() { R = 0, G = 120, B = 215, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonShadow:
                            rgbColor = new RgbaColor() { R = 160, G = 160, B = 160, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonText:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.CaptionText:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientActiveCaption:
                            rgbColor = new RgbaColor() { R = 185, G = 209, B = 234, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientInactiveCaption:
                            rgbColor = new RgbaColor() { R = 215, G = 228, B = 242, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.GrayText:
                            rgbColor = new RgbaColor() { R = 109, G = 109, B = 109, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.Highlight:
                            rgbColor = new RgbaColor() { R = 0, G = 120, B = 215, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.HighlightText:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.HotLight:
                            rgbColor = new RgbaColor() { R = 255, G = 165, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveBorder:
                            rgbColor = new RgbaColor() { R = 244, G = 247, B = 252, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaption:
                            rgbColor = new RgbaColor() { R = 191, G = 205, B = 219, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaptionText:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoBack:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 225, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoText:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.Menu:
                            rgbColor = new RgbaColor() { R = 240, G = 240, B = 240, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuBar:
                            rgbColor = new RgbaColor() { R = 240, G = 240, B = 240, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuHighlight:
                            rgbColor = new RgbaColor() { R = 0, G = 120, B = 215, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuText:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ScrollBar:
                            rgbColor = new RgbaColor() { R = 200, G = 200, B = 200, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDDarkShadow:
                            rgbColor = new RgbaColor() { R = 160, G = 160, B = 160, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDLight:
                            rgbColor = new RgbaColor() { R = 227, G = 227, B = 227, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.Window:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowFrame:
                            rgbColor = new RgbaColor() { R = 100, G = 100, B = 100, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        default:
                            return string.Empty;
                    };
                }
                else if (color.PresetColor != null && color.PresetColor.Val != null && color.PresetColor.Val.HasValue)
                {
                    switch (color.PresetColor.Val.Value)
                    {
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.AliceBlue:
                            rgbColor = new RgbaColor() { R = 240, G = 248, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.AntiqueWhite:
                            rgbColor = new RgbaColor() { R = 250, G = 235, B = 215, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Aqua:
                            rgbColor = new RgbaColor() { R = 0, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Aquamarine:
                            rgbColor = new RgbaColor() { R = 127, G = 255, B = 212, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Azure:
                            rgbColor = new RgbaColor() { R = 240, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Beige:
                            rgbColor = new RgbaColor() { R = 245, G = 245, B = 220, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Bisque:
                            rgbColor = new RgbaColor() { R = 255, G = 228, B = 196, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Black:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.BlanchedAlmond:
                            rgbColor = new RgbaColor() { R = 255, G = 235, B = 205, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Blue:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.BlueViolet:
                            rgbColor = new RgbaColor() { R = 138, G = 43, B = 226, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Brown:
                            rgbColor = new RgbaColor() { R = 165, G = 42, B = 42, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.BurlyWood:
                            rgbColor = new RgbaColor() { R = 222, G = 184, B = 135, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.CadetBlue:
                            rgbColor = new RgbaColor() { R = 95, G = 158, B = 160, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Chartreuse:
                            rgbColor = new RgbaColor() { R = 127, G = 255, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Chocolate:
                            rgbColor = new RgbaColor() { R = 210, G = 105, B = 30, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Coral:
                            rgbColor = new RgbaColor() { R = 255, G = 127, B = 80, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.CornflowerBlue:
                            rgbColor = new RgbaColor() { R = 100, G = 149, B = 237, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Cornsilk:
                            rgbColor = new RgbaColor() { R = 255, G = 248, B = 220, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Crimson:
                            rgbColor = new RgbaColor() { R = 220, G = 20, B = 60, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Cyan:
                            rgbColor = new RgbaColor() { R = 0, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan:
                            rgbColor = new RgbaColor() { R = 0, G = 139, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod:
                            rgbColor = new RgbaColor() { R = 184, G = 134, B = 11, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray:
                            rgbColor = new RgbaColor() { R = 169, G = 169, B = 169, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen:
                            rgbColor = new RgbaColor() { R = 0, G = 100, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki:
                            rgbColor = new RgbaColor() { R = 189, G = 183, B = 107, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta:
                            rgbColor = new RgbaColor() { R = 139, G = 0, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen:
                            rgbColor = new RgbaColor() { R = 85, G = 107, B = 47, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange:
                            rgbColor = new RgbaColor() { R = 255, G = 140, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid:
                            rgbColor = new RgbaColor() { R = 153, G = 50, B = 204, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed:
                            rgbColor = new RgbaColor() { R = 139, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon:
                            rgbColor = new RgbaColor() { R = 233, G = 150, B = 122, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen:
                            rgbColor = new RgbaColor() { R = 143, G = 188, B = 143, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue:
                            rgbColor = new RgbaColor() { R = 72, G = 61, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray:
                            rgbColor = new RgbaColor() { R = 47, G = 79, B = 79, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise:
                            rgbColor = new RgbaColor() { R = 0, G = 206, B = 209, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet:
                            rgbColor = new RgbaColor() { R = 148, G = 0, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepPink:
                            rgbColor = new RgbaColor() { R = 255, G = 20, B = 147, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepSkyBlue:
                            rgbColor = new RgbaColor() { R = 0, G = 191, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGray:
                            rgbColor = new RgbaColor() { R = 105, G = 105, B = 105, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DodgerBlue:
                            rgbColor = new RgbaColor() { R = 30, G = 144, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Firebrick:
                            rgbColor = new RgbaColor() { R = 178, G = 34, B = 34, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.FloralWhite:
                            rgbColor = new RgbaColor() { R = 255, G = 250, B = 240, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.ForestGreen:
                            rgbColor = new RgbaColor() { R = 34, G = 139, B = 34, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Fuchsia:
                            rgbColor = new RgbaColor() { R = 255, G = 0, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Gainsboro:
                            rgbColor = new RgbaColor() { R = 220, G = 220, B = 220, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.GhostWhite:
                            rgbColor = new RgbaColor() { R = 248, G = 248, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Gold:
                            rgbColor = new RgbaColor() { R = 255, G = 215, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Goldenrod:
                            rgbColor = new RgbaColor() { R = 218, G = 165, B = 32, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Gray:
                            rgbColor = new RgbaColor() { R = 128, G = 128, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Green:
                            rgbColor = new RgbaColor() { R = 0, G = 128, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.GreenYellow:
                            rgbColor = new RgbaColor() { R = 173, G = 255, B = 47, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Honeydew:
                            rgbColor = new RgbaColor() { R = 240, G = 255, B = 240, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.HotPink:
                            rgbColor = new RgbaColor() { R = 255, G = 105, B = 180, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.IndianRed:
                            rgbColor = new RgbaColor() { R = 205, G = 92, B = 92, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Indigo:
                            rgbColor = new RgbaColor() { R = 75, G = 0, B = 130, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Ivory:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 240, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Khaki:
                            rgbColor = new RgbaColor() { R = 240, G = 230, B = 140, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Lavender:
                            rgbColor = new RgbaColor() { R = 230, G = 230, B = 250, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LavenderBlush:
                            rgbColor = new RgbaColor() { R = 255, G = 240, B = 245, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LawnGreen:
                            rgbColor = new RgbaColor() { R = 124, G = 252, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LemonChiffon:
                            rgbColor = new RgbaColor() { R = 255, G = 250, B = 205, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue:
                            rgbColor = new RgbaColor() { R = 173, G = 216, B = 230, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral:
                            rgbColor = new RgbaColor() { R = 240, G = 128, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan:
                            rgbColor = new RgbaColor() { R = 224, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow:
                            rgbColor = new RgbaColor() { R = 250, G = 250, B = 210, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray:
                            rgbColor = new RgbaColor() { R = 211, G = 211, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen:
                            rgbColor = new RgbaColor() { R = 144, G = 238, B = 144, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink:
                            rgbColor = new RgbaColor() { R = 255, G = 182, B = 193, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon:
                            rgbColor = new RgbaColor() { R = 255, G = 160, B = 122, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen:
                            rgbColor = new RgbaColor() { R = 32, G = 178, B = 170, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue:
                            rgbColor = new RgbaColor() { R = 135, G = 206, B = 250, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray:
                            rgbColor = new RgbaColor() { R = 119, G = 136, B = 153, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue:
                            rgbColor = new RgbaColor() { R = 176, G = 196, B = 222, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 224, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Lime:
                            rgbColor = new RgbaColor() { R = 0, G = 255, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LimeGreen:
                            rgbColor = new RgbaColor() { R = 50, G = 205, B = 50, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Linen:
                            rgbColor = new RgbaColor() { R = 250, G = 240, B = 230, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Magenta:
                            rgbColor = new RgbaColor() { R = 255, G = 0, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Maroon:
                            rgbColor = new RgbaColor() { R = 128, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MedAquamarine:
                            rgbColor = new RgbaColor() { R = 102, G = 205, B = 170, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 205, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid:
                            rgbColor = new RgbaColor() { R = 186, G = 85, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple:
                            rgbColor = new RgbaColor() { R = 147, G = 112, B = 219, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen:
                            rgbColor = new RgbaColor() { R = 60, G = 179, B = 113, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue:
                            rgbColor = new RgbaColor() { R = 123, G = 104, B = 238, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen:
                            rgbColor = new RgbaColor() { R = 0, G = 250, B = 154, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise:
                            rgbColor = new RgbaColor() { R = 72, G = 209, B = 204, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed:
                            rgbColor = new RgbaColor() { R = 199, G = 21, B = 133, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MidnightBlue:
                            rgbColor = new RgbaColor() { R = 25, G = 25, B = 112, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MintCream:
                            rgbColor = new RgbaColor() { R = 245, G = 255, B = 250, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MistyRose:
                            rgbColor = new RgbaColor() { R = 255, G = 228, B = 225, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Moccasin:
                            rgbColor = new RgbaColor() { R = 255, G = 228, B = 181, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.NavajoWhite:
                            rgbColor = new RgbaColor() { R = 255, G = 222, B = 173, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Navy:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.OldLace:
                            rgbColor = new RgbaColor() { R = 253, G = 245, B = 230, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Olive:
                            rgbColor = new RgbaColor() { R = 128, G = 128, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.OliveDrab:
                            rgbColor = new RgbaColor() { R = 107, G = 142, B = 35, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Orange:
                            rgbColor = new RgbaColor() { R = 255, G = 165, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.OrangeRed:
                            rgbColor = new RgbaColor() { R = 255, G = 69, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Orchid:
                            rgbColor = new RgbaColor() { R = 218, G = 112, B = 214, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGoldenrod:
                            rgbColor = new RgbaColor() { R = 238, G = 232, B = 170, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGreen:
                            rgbColor = new RgbaColor() { R = 152, G = 251, B = 152, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleTurquoise:
                            rgbColor = new RgbaColor() { R = 175, G = 238, B = 238, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleVioletRed:
                            rgbColor = new RgbaColor() { R = 219, G = 112, B = 147, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PapayaWhip:
                            rgbColor = new RgbaColor() { R = 255, G = 239, B = 213, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PeachPuff:
                            rgbColor = new RgbaColor() { R = 255, G = 218, B = 185, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Peru:
                            rgbColor = new RgbaColor() { R = 205, G = 133, B = 63, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Pink:
                            rgbColor = new RgbaColor() { R = 255, G = 192, B = 203, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Plum:
                            rgbColor = new RgbaColor() { R = 221, G = 160, B = 221, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PowderBlue:
                            rgbColor = new RgbaColor() { R = 176, G = 224, B = 230, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Purple:
                            rgbColor = new RgbaColor() { R = 128, G = 0, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Red:
                            rgbColor = new RgbaColor() { R = 255, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.RosyBrown:
                            rgbColor = new RgbaColor() { R = 188, G = 143, B = 143, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.RoyalBlue:
                            rgbColor = new RgbaColor() { R = 65, G = 105, B = 225, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SaddleBrown:
                            rgbColor = new RgbaColor() { R = 139, G = 69, B = 19, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Salmon:
                            rgbColor = new RgbaColor() { R = 250, G = 128, B = 114, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SandyBrown:
                            rgbColor = new RgbaColor() { R = 244, G = 164, B = 96, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaGreen:
                            rgbColor = new RgbaColor() { R = 46, G = 139, B = 87, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaShell:
                            rgbColor = new RgbaColor() { R = 255, G = 245, B = 238, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Sienna:
                            rgbColor = new RgbaColor() { R = 160, G = 82, B = 45, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Silver:
                            rgbColor = new RgbaColor() { R = 192, G = 192, B = 192, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SkyBlue:
                            rgbColor = new RgbaColor() { R = 135, G = 206, B = 235, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateBlue:
                            rgbColor = new RgbaColor() { R = 106, G = 90, B = 205, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGray:
                            rgbColor = new RgbaColor() { R = 112, G = 128, B = 144, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Snow:
                            rgbColor = new RgbaColor() { R = 255, G = 250, B = 250, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SpringGreen:
                            rgbColor = new RgbaColor() { R = 0, G = 255, B = 127, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SteelBlue:
                            rgbColor = new RgbaColor() { R = 70, G = 130, B = 180, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Tan:
                            rgbColor = new RgbaColor() { R = 210, G = 180, B = 140, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Teal:
                            rgbColor = new RgbaColor() { R = 0, G = 128, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Thistle:
                            rgbColor = new RgbaColor() { R = 216, G = 191, B = 216, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Tomato:
                            rgbColor = new RgbaColor() { R = 255, G = 99, B = 71, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Turquoise:
                            rgbColor = new RgbaColor() { R = 64, G = 224, B = 208, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Violet:
                            rgbColor = new RgbaColor() { R = 238, G = 130, B = 238, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Wheat:
                            rgbColor = new RgbaColor() { R = 245, G = 222, B = 179, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.White:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.WhiteSmoke:
                            rgbColor = new RgbaColor() { R = 245, G = 245, B = 245, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Yellow:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.YellowGreen:
                            rgbColor = new RgbaColor() { R = 154, G = 205, B = 50, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue2010:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan2010:
                            rgbColor = new RgbaColor() { R = 0, G = 139, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod2010:
                            rgbColor = new RgbaColor() { R = 184, G = 134, B = 11, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray2010:
                            rgbColor = new RgbaColor() { R = 169, G = 169, B = 169, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey2010:
                            rgbColor = new RgbaColor() { R = 169, G = 169, B = 169, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen2010:
                            rgbColor = new RgbaColor() { R = 0, G = 100, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki2010:
                            rgbColor = new RgbaColor() { R = 189, G = 183, B = 107, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta2010:
                            rgbColor = new RgbaColor() { R = 139, G = 0, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen2010:
                            rgbColor = new RgbaColor() { R = 85, G = 107, B = 47, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange2010:
                            rgbColor = new RgbaColor() { R = 255, G = 140, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid2010:
                            rgbColor = new RgbaColor() { R = 153, G = 50, B = 204, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed2010:
                            rgbColor = new RgbaColor() { R = 139, G = 0, B = 0, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon2010:
                            rgbColor = new RgbaColor() { R = 233, G = 150, B = 122, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen2010:
                            rgbColor = new RgbaColor() { R = 143, G = 188, B = 143, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue2010:
                            rgbColor = new RgbaColor() { R = 72, G = 61, B = 139, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray2010:
                            rgbColor = new RgbaColor() { R = 47, G = 79, B = 79, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey2010:
                            rgbColor = new RgbaColor() { R = 47, G = 79, B = 79, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise2010:
                            rgbColor = new RgbaColor() { R = 0, G = 206, B = 209, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet2010:
                            rgbColor = new RgbaColor() { R = 148, G = 0, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue2010:
                            rgbColor = new RgbaColor() { R = 173, G = 216, B = 230, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral2010:
                            rgbColor = new RgbaColor() { R = 240, G = 128, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan2010:
                            rgbColor = new RgbaColor() { R = 224, G = 255, B = 255, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow2010:
                            rgbColor = new RgbaColor() { R = 250, G = 250, B = 210, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray2010:
                            rgbColor = new RgbaColor() { R = 211, G = 211, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey2010:
                            rgbColor = new RgbaColor() { R = 211, G = 211, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen2010:
                            rgbColor = new RgbaColor() { R = 144, G = 238, B = 144, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink2010:
                            rgbColor = new RgbaColor() { R = 255, G = 182, B = 193, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon2010:
                            rgbColor = new RgbaColor() { R = 255, G = 160, B = 122, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen2010:
                            rgbColor = new RgbaColor() { R = 32, G = 178, B = 170, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue2010:
                            rgbColor = new RgbaColor() { R = 135, G = 206, B = 250, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray2010:
                            rgbColor = new RgbaColor() { R = 119, G = 136, B = 153, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey2010:
                            rgbColor = new RgbaColor() { R = 119, G = 136, B = 153, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue2010:
                            rgbColor = new RgbaColor() { R = 176, G = 196, B = 222, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow2010:
                            rgbColor = new RgbaColor() { R = 255, G = 255, B = 224, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumAquamarine2010:
                            rgbColor = new RgbaColor() { R = 102, G = 205, B = 170, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue2010:
                            rgbColor = new RgbaColor() { R = 0, G = 0, B = 205, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid2010:
                            rgbColor = new RgbaColor() { R = 186, G = 85, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple2010:
                            rgbColor = new RgbaColor() { R = 147, G = 112, B = 219, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen2010:
                            rgbColor = new RgbaColor() { R = 60, G = 179, B = 113, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue2010:
                            rgbColor = new RgbaColor() { R = 123, G = 104, B = 238, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen2010:
                            rgbColor = new RgbaColor() { R = 0, G = 250, B = 154, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise2010:
                            rgbColor = new RgbaColor() { R = 72, G = 209, B = 204, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed2010:
                            rgbColor = new RgbaColor() { R = 199, G = 21, B = 133, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey:
                            rgbColor = new RgbaColor() { R = 169, G = 169, B = 169, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGrey:
                            rgbColor = new RgbaColor() { R = 105, G = 105, B = 105, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey:
                            rgbColor = new RgbaColor() { R = 47, G = 79, B = 79, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Grey:
                            rgbColor = new RgbaColor() { R = 128, G = 128, B = 128, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey:
                            rgbColor = new RgbaColor() { R = 211, G = 211, B = 211, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey:
                            rgbColor = new RgbaColor() { R = 119, G = 136, B = 153, A = 1 };
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGrey:
                            rgbColor = new RgbaColor() { R = 112, G = 128, B = 144, A = 1 };
                            break;
                        default:
                            return string.Empty;
                    };
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }

            if (type.Tint != null && type.Tint.HasValue)
            {
                int r = rgbColor.R;
                int g = rgbColor.G;
                int b = rgbColor.B;
                RgbToHls(r, g, b, out double h, out double l, out double s);

                if (type.Tint.Value < 0)
                {
                    HlsToRgb(h, l * (1 + type.Tint.Value), s, out r, out g, out b);
                }
                else
                {
                    HlsToRgb(h, l * (1 - type.Tint.Value) + 1 - 1 * (1 - type.Tint.Value), s, out r, out g, out b);
                }

                rgbColor.R = r;
                rgbColor.G = g;
                rgbColor.B = b;
            }

            if (rgbColor.A >= 1)
            {
                return $"rgb({rgbColor.R}, {rgbColor.G}, {rgbColor.B})";
            }
            else
            {
                return $"rgba({rgbColor.R}, {rgbColor.G}, {rgbColor.B}, {rgbColor.A})";
            }
        }

        private static RgbaColor HexToRgba(string hex)
        {
            string hexTrimmed = hex.Replace("#", string.Empty);
            if (hexTrimmed.Length < 6)
            {
                return new RgbaColor() { R = 0, G = 0, B = 0, A = 0 };
            }
            else
            {
                return new RgbaColor()
                {
                    R = Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length >= 8 ? 2 : 0, 2), 16),
                    G = Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length >= 8 ? 4 : 2, 2), 16),
                    B = Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length >= 8 ? 6 : 4, 2), 16),
                    A = hexTrimmed.Length >= 8 ? Convert.ToInt32(hexTrimmed.Substring(0, 2), 16) / 255.0 : 1
                };
            }
        }

        private static void RgbToHls(int r, int g, int b, out double h, out double l, out double s)
        {
            double red = r / 255.0;
            double green = g / 255.0;
            double blue = b / 255.0;

            double max = Math.Max(red, Math.Max(green, blue));
            double min = Math.Min(red, Math.Min(green, blue));

            l = (max + min) / 2;

            double difference = max - min;
            if (Math.Abs(difference) < 0.00001)
            {
                s = 0;
                h = 0;
            }
            else
            {
                s = l <= 0.5 ? difference / (max + min) : difference / (2 - max - min);

                double distanceRed = (max - red) / difference;
                double distanceGreen = (max - green) / difference;
                double distanceBlue = (max - blue) / difference;

                h = (red == max ? distanceBlue - distanceGreen : (green == max ? 2 + distanceRed - distanceBlue : 4 + distanceGreen - distanceRed)) * 60;
                h += h < 0 ? 360 : 0;
            }
        }

        private static void HlsToRgb(double h, double l, double s, out int r, out int g, out int b)
        {
            double p2 = l <= 0.5 ? l * (1 + s) : l + s - l * s;
            double p1 = 2 * l - p2;

            r = (int)((s == 0 ? l : QqhToRgb(p1, p2, h + 120)) * 255.0);
            g = (int)((s == 0 ? l : QqhToRgb(p1, p2, h)) * 255.0);
            b = (int)((s == 0 ? l : QqhToRgb(p1, p2, h - 120)) * 255.0);
        }

        private static double QqhToRgb(double q1, double q2, double hue)
        {
            hue -= hue > 360 ? 360 : 0;
            hue += hue < 0 ? 360 : 0;
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
            else
            {
                return q1;
            }
        }

        private static Dictionary<string, string> FontToHtml(WorkbookPart workbook, ColorType color, FontSize fontSize, Bold bold, Italic italic, Strike strike, Underline underline)
        {
            Dictionary<string, string> htmlStyle = new Dictionary<string, string>();

            if (color != null)
            {
                string htmlColor = ColorTypeToHtml(workbook, color);
                if (!string.IsNullOrEmpty(htmlColor))
                {
                    htmlStyle.Add("color", htmlColor);
                }
            }
            if (fontSize != null && fontSize.Val != null && fontSize.Val.HasValue)
            {
                htmlStyle.Add("font-size", $"{fontSize.Val.Value * 96 / 72}px");
            }
            if (bold != null)
            {
                htmlStyle.Add("font-weight", bold.Val == null || (bold.Val.HasValue && bold.Val.Value) ? "bold" : "normal");
            }
            if (italic != null)
            {
                htmlStyle.Add("font-style", italic.Val == null || (italic.Val.HasValue && italic.Val.Value) ? "italic" : "normal");
            }
            string textDecoraion = string.Empty;
            if (strike != null)
            {
                textDecoraion += strike.Val == null || (strike.Val.HasValue && strike.Val.Value) ? " line-through" : " none";
            }
            if (underline != null && underline.Val != null && underline.Val.HasValue)
            {
                switch (underline.Val.Value)
                {
                    case UnderlineValues.Single:
                        textDecoraion += " underline";
                        break;
                    case UnderlineValues.SingleAccounting:
                        textDecoraion += " underline";
                        break;
                    case UnderlineValues.Double:
                        textDecoraion += " underline double";
                        break;
                    case UnderlineValues.DoubleAccounting:
                        textDecoraion += " underline double";
                        break;
                    default:
                        textDecoraion += " underline";
                        break;
                }
            }
            if (!string.IsNullOrEmpty(textDecoraion))
            {
                htmlStyle.Add("text-decoration", textDecoraion.Trim());
            }

            return htmlStyle;
        }

        private static void BorderPropertiesToHtml(WorkbookPart workbook, BorderPropertiesType border, ref string width, ref string style, ref string color)
        {
            if (border == null)
            {
                return;
            }

            if (border.Style != null && border.Style.HasValue)
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

            if (border.Color != null)
            {
                string value = ColorTypeToHtml(workbook, border.Color);
                if (!string.IsNullOrEmpty(value))
                {
                    color = value;
                }
            }
        }

        private static string GetGeneralAlignment(Cell cell, string cellValue)
        {
            if (cell != null && cell.DataType != null && cell.DataType.HasValue)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.Error:
                        return "center";
                    case CellValues.Boolean:
                        return "center";
                    case CellValues.Number:
                        return "right";
                    default:
                        return "left";
                }
            }
            else
            {
                return !string.IsNullOrEmpty(cellValue) && double.TryParse(cellValue, out double _) ? "right" : "left";
            }
        }

        private static void DrawingsToHtml(WorksheetPart worksheet, OpenXmlCompositeElement anchor, StreamWriter writer, double left, double top, double width, double height, bool isPictureAllowed)
        {
            if (anchor == null)
            {
                return;
            }

            List<DrawingInfo> drawings = new List<DrawingInfo>();

            if (isPictureAllowed)
            {
                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>())
                {
                    if (picture.BlipFill == null || picture.BlipFill.Blip == null)
                    {
                        continue;
                    }

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties properties = null;
                    if (picture.NonVisualPictureProperties != null)
                    {
                        properties = picture.NonVisualPictureProperties.NonVisualDrawingProperties;
                    }

                    if (picture.BlipFill.Blip.Embed != null && picture.BlipFill.Blip.Embed.HasValue)
                    {
                        ImagePart imagePart = worksheet.DrawingsPart.GetPartById(picture.BlipFill.Blip.Embed.Value) as ImagePart;

                        Stream imageStream = imagePart.GetStream();
                        if (imageStream.CanSeek)
                        {
                            imageStream.Seek(0, SeekOrigin.Begin);
                        }
                        byte[] data = new byte[imageStream.Length];
                        imageStream.Read(data, 0, (int)imageStream.Length);
                        string base64 = Convert.ToBase64String(data, Base64FormattingOptions.None);

                        drawings.Add(new DrawingInfo()
                        {
                            Prefix = $"<img src=\"data:{imagePart.ContentType};base64,{base64}\"{(properties != null && properties.Description != null && properties.Description.HasValue ? " alt=\"" + properties.Description.Value + "\"" : string.Empty)}",
                            Postfix = "/>",
                            IsHidden = properties != null && properties.Hidden != null && properties.Hidden.HasValue && properties.Hidden.Value,
                            Left = left,
                            Top = top,
                            Width = width,
                            Height = height,
                            ShapeProperties = picture.ShapeProperties
                        });
                    }
                }
            }

            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>())
            {
                string text = shape.TextBody != null ? shape.TextBody.InnerText : string.Empty;

                DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties properties = null;
                if (shape.NonVisualShapeProperties != null)
                {
                    properties = shape.NonVisualShapeProperties.NonVisualDrawingProperties;
                }

                //TODO: shape styles
                drawings.Add(new DrawingInfo()
                {
                    Prefix = "<p",
                    Postfix = $">{text}</p>",
                    IsHidden = properties != null && properties.Hidden != null && properties.Hidden.HasValue && properties.Hidden.Value,
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height,
                    ShapeProperties = shape.ShapeProperties
                });
            }

            foreach (DrawingInfo drawingInfo in drawings)
            {
                double leftActual = drawingInfo.Left;
                double topActual = drawingInfo.Top;
                double widthActual = drawingInfo.Width;
                double heightActual = drawingInfo.Height;
                double rotation = 0;

                if (drawingInfo.ShapeProperties != null && drawingInfo.ShapeProperties.Transform2D != null)
                {
                    if (drawingInfo.ShapeProperties.Transform2D.Offset != null)
                    {
                        if (!double.IsNaN(leftActual))
                        {
                            leftActual += drawingInfo.ShapeProperties.Transform2D.Offset.X != null && drawingInfo.ShapeProperties.Transform2D.Offset.X.HasValue ? drawingInfo.ShapeProperties.Transform2D.Offset.X.Value / 914400 * 96 : 0;
                        }
                        if (!double.IsNaN(topActual))
                        {
                            topActual += drawingInfo.ShapeProperties.Transform2D.Offset.Y != null && drawingInfo.ShapeProperties.Transform2D.Offset.Y.HasValue ? drawingInfo.ShapeProperties.Transform2D.Offset.Y.Value / 914400 * 96 : 0;
                        }
                    }
                    if (drawingInfo.ShapeProperties.Transform2D.Extents != null)
                    {
                        if (!double.IsNaN(widthActual))
                        {
                            widthActual += drawingInfo.ShapeProperties.Transform2D.Extents.Cx != null && drawingInfo.ShapeProperties.Transform2D.Extents.Cx.HasValue ? drawingInfo.ShapeProperties.Transform2D.Extents.Cx.Value / 914400 * 96 : 0;
                        }
                        if (!double.IsNaN(heightActual))
                        {
                            heightActual += drawingInfo.ShapeProperties.Transform2D.Extents.Cy != null && drawingInfo.ShapeProperties.Transform2D.Extents.Cy.HasValue ? drawingInfo.ShapeProperties.Transform2D.Extents.Cy.Value / 914400 * 96 : 0;
                        }
                    }
                    if (drawingInfo.ShapeProperties.Transform2D.Rotation != null && drawingInfo.ShapeProperties.Transform2D.Rotation.HasValue)
                    {
                        rotation = drawingInfo.ShapeProperties.Transform2D.Rotation.Value;
                    }
                }

                writer.Write($"\n{new string(' ', 8)}{drawingInfo.Prefix} style=\"position: absolute; left: {(!double.IsNaN(leftActual) ? leftActual : 0)}px; top: {(!double.IsNaN(topActual) ? topActual : 0)}px; width: {(!double.IsNaN(widthActual) ? widthActual + "px" : "auto")}; height: {(!double.IsNaN(heightActual) ? heightActual + "px" : "auto")}px;{(rotation != 0 ? $" transform: rotate(-{rotation}deg);" : string.Empty)}{(drawingInfo.IsHidden ? " visibility: hidden;" : string.Empty)}\"{drawingInfo.Postfix}");
            }
        }

        #endregion

        #region Private Structures

        private struct RgbaColor
        {
            public int R { get; set; }
            public int G { get; set; }
            public int B { get; set; }
            public double A { get; set; }
        }
        private struct MergeCellInfo
        {
            public int FromColumn { get; set; }
            public int FromRow { get; set; }
            public int ToColumn { get; set; }
            public int ToRow { get; set; }
            public int ColumnSpanned { get; set; }
            public int RowSpanned { get; set; }
        }

        private struct DrawingInfo
        {
            public string Prefix { get; set; }
            public string Postfix { get; set; }
            public bool IsHidden { get; set; }
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties ShapeProperties { get; set; }
        }

        #endregion

        #region Private Fields

        private static readonly Regex regexNumbers = new Regex(@"\d+", RegexOptions.Compiled);
        private static readonly Regex regexLetters = new Regex("[A-Za-z]+", RegexOptions.Compiled);

        #endregion
    }

    /// <summary>
    /// The configurations of the Xlsx to Html converter.
    /// </summary>
    public class ConverterConfig
    {
        public const string DefaultErrorMessage = "Error, unable to convert XLSX file. The file is either already open in another program (please close it first) or contains corrupted data.";
        public const string DefaultPresetStyles = @"body {
            margin: 0;
            padding: 0;
            width: 100%;
        }

        h5 {
            font-size: 20px;
            font-weight: bold;
            font-family: monospace;
            text-align: center;
            width: fit-content;
            margin: 10px auto;
            border-bottom-width: 4px;
            border-bottom-style: solid;
            border-bottom-color: transparent;
            padding-bottom: 3px;
        }

        table {
            width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
        }

        td {
            text-align: left;
            vertical-align: bottom;
            padding: 0px;
            color: black;
            background-color: transparent;
            border-width: 1px;
            border-style: solid;
            border-color: lightgray;
            border-collapse: collapse;
            white-space: nowrap;
            overflow: hidden;
        }";

        /// <summary>
        /// Initializes a new instance of the <see cref="ConverterConfig"/> class.
        /// </summary>
        public ConverterConfig()
        {
            this.PageTitle = "Title";
            this.PresetStyles = DefaultPresetStyles;
            this.ErrorMessage = DefaultErrorMessage;
            this.Encoding = System.Text.Encoding.UTF8;
            this.ConvertStyle = true;
            this.ConvertSize = true;
            this.ConvertPicture = true;
            this.ConvertSheetNameTitle = true;
            this.ConvertHiddenSheet = false;
            this.ConvertFirstSheetOnly = false;
            this.ConvertHtmlBodyOnly = false;
        }

        #region Public Fields

        /// <summary>
        /// Gets or sets the Html page title.
        /// </summary>
        public string PageTitle { get; set; }

        /// <summary>
        /// Gets or sets the preset CSS style in Html.
        /// </summary>
        public string PresetStyles { get; set; }

        /// <summary>
        /// Gets or sets the error message that will show when convert failed. Text "{EXCEPTION}" will be replaced by the exception message.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets the encoding to use when writing the Html string.
        /// </summary>
        public System.Text.Encoding Encoding { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx styles into Html styles or not.
        /// </summary>
        public bool ConvertStyle { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx cell sizes into Html table cell sizes or not.
        /// </summary>
        public bool ConvertSize { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx pictures into Html pictures or not.
        /// </summary>
        public bool ConvertPicture { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx sheet names to titles or not.
        /// </summary>
        public bool ConvertSheetNameTitle { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx hidden sheets or not.
        /// </summary>
        public bool ConvertHiddenSheet { get; set; }

        /// <summary>
        /// Gets or sets whether to only convert the first Xlsx sheet or not.
        /// </summary>
        public bool ConvertFirstSheetOnly { get; set; }

        /// <summary>
        /// Gets or sets whether to only convert into the body tag of Html or not.
        /// </summary>
        public bool ConvertHtmlBodyOnly { get; set; }

        /// <summary>
        /// Gets a new instance of <see cref="ConverterConfig">ConverterConfig</see> with default settings.
        /// </summary>
        public static ConverterConfig DefaultSettings { get { return new ConverterConfig(); } }

        #endregion
    }

    /// <summary>
    /// The progress callback event arguments class of the Xlsx to Html converter.
    /// </summary>
    public class ConverterProgressCallbackEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConverterProgressCallbackEventArgs"/> class with specific current sheet and total sheets.
        /// </summary>
        public ConverterProgressCallbackEventArgs(int currentSheet, int totalSheets, int currentRow, int totalRows)
        {
            this.CurrentSheet = currentSheet;
            this.TotalSheets = totalSheets;
            this.CurrentRow = currentRow;
            this.TotalRows = totalRows;
        }

        #region Public Fields

        /// <summary>
        /// Gets the current progress in percentage.
        /// </summary>
        public double ProgressPercent
        {
            get
            {
                return Math.Max(0, Math.Min(100, (double)(CurrentSheet - 1) / TotalSheets * 100 + (double)CurrentRow / TotalRows * (100 / (double)TotalSheets)));
            }
        }

        /// <summary>
        /// Gets the 1-indexed id of the current sheet.
        /// </summary>
        public int CurrentSheet { get; private set; }

        /// <summary>
        /// Gets the total number of sheets in the Xlsx file.
        /// </summary>
        public int TotalSheets { get; private set; }

        /// <summary>
        /// Gets the 1-indexed number of the current row in the current sheet.
        /// </summary>
        public int CurrentRow { get; private set; }

        /// <summary>
        /// Gets the total number of sheets in the current sheet.
        /// </summary>
        public int TotalRows { get; private set; }

        #endregion
    }
}
