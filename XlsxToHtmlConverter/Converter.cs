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
                    for (int stylesheetDifferentialFormatIndex = 0; stylesheetDifferentialFormatIndex < htmlStyleDifferentialFormats.Length; stylesheetDifferentialFormatIndex++)
                    {
                        if (stylesheet.DifferentialFormats.ChildElements[stylesheetDifferentialFormatIndex] is DifferentialFormat differentialFormat)
                        {
                            htmlStyleDifferentialFormats[stylesheetDifferentialFormatIndex] = new Tuple<Dictionary<string, string>, string>(CellFormatToHtml(workbook, differentialFormat.Fill, differentialFormat.Font, differentialFormat.Border, differentialFormat.Alignment, out string cellValueContainer), cellValueContainer);
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
                                foreach (OpenXmlElement element in sharedString.Elements())
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
                                        if (config.ConvertStyles && run.RunProperties is RunProperties runProperties)
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
                    foreach (Sheet sheet in sheets)
                    {
                        sheetIndex++;
                        if ((config.ConvertFirstSheetOnly && sheetIndex > 1) || (!config.ConvertHiddenSheets && sheet.State != null && sheet.State.HasValue && sheet.State.Value != SheetStateValues.Visible) || !(workbook.GetPartById(sheet.Id) is WorksheetPart worksheetPart))
                        {
                            continue;
                        }

                        Worksheet worksheet = worksheetPart.Worksheet;

                        if (config.ConvertSheetTitles)
                        {
                            string tabColor = worksheet.SheetProperties != null && worksheet.SheetProperties.TabColor != null ? ColorTypeToHtml(workbook, worksheet.SheetProperties.TabColor) : string.Empty;
                            writer.Write($"\n{new string(' ', 4)}<h5{(!string.IsNullOrEmpty(tabColor) ? $" style=\"border-bottom-color: {tabColor};\"" : string.Empty)}>{(sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : "Untitled")}</h5>");
                        }

                        writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                        writer.Write($"\n{new string(' ', 8)}<table>");

                        bool isMergeCellsContained = false;
                        List<MergeCellInfo> mergeCells = new List<MergeCellInfo>();
                        if (worksheet.Descendants<MergeCells>().FirstOrDefault() is MergeCells mergeCellsGroup)
                        {
                            isMergeCellsContained = true;
                            foreach (MergeCell mergeCell in mergeCellsGroup.Cast<MergeCell>())
                            {
                                if (mergeCell.Reference == null || !mergeCell.Reference.HasValue)
                                {
                                    continue;
                                }

                                GetReferenceRange(mergeCell.Reference.Value.Split(':'), out int mergeCellFromColumn, out int mergeCellFromRow, out int mergeCellToColumn, out int mergeCellToRow);
                                mergeCells.Add(new MergeCellInfo()
                                {
                                    FromColumn = mergeCellFromColumn,
                                    ToColumn = mergeCellToColumn,
                                    FromRow = mergeCellFromRow,
                                    ToRow = mergeCellToRow,
                                    ColumnSpanned = mergeCellToColumn - mergeCellFromColumn + 1,
                                    RowSpanned = mergeCellToRow - mergeCellFromRow + 1
                                });
                            }
                        }

                        int[] sheetDimension = new int[4] { 1, 1, 1, 1 };
                        if (worksheet.SheetDimension != null && worksheet.SheetDimension.Reference != null && worksheet.SheetDimension.Reference.HasValue)
                        {
                            GetReferenceRange(worksheet.SheetDimension.Reference.Value.Split(':'), out sheetDimension[0], out sheetDimension[1], out sheetDimension[2], out sheetDimension[3]);
                        }
                        else
                        {
                            foreach (Cell cell in worksheet.Descendants<Row>().SelectMany(x => x.Descendants<Cell>()))
                            {
                                if (cell.CellReference != null && cell.CellReference.HasValue)
                                {
                                    sheetDimension[2] = Math.Max(sheetDimension[2], GetColumnIndex(cell.CellReference.Value));
                                    sheetDimension[3] = Math.Max(sheetDimension[3], GetRowIndex(cell.CellReference.Value));
                                }
                            }
                        }

                        double columnWidthDefault = worksheet.SheetFormatProperties != null && worksheet.SheetFormatProperties.DefaultColumnWidth != null && worksheet.SheetFormatProperties.DefaultColumnWidth.HasValue ? worksheet.SheetFormatProperties.DefaultColumnWidth.Value : (worksheet.SheetFormatProperties != null && worksheet.SheetFormatProperties.BaseColumnWidth != null && worksheet.SheetFormatProperties.BaseColumnWidth.HasValue ? worksheet.SheetFormatProperties.BaseColumnWidth.Value : double.NaN);
                        double rowHeightDefault = worksheet.SheetFormatProperties != null && worksheet.SheetFormatProperties.DefaultRowHeight != null && worksheet.SheetFormatProperties.DefaultRowHeight.HasValue ? worksheet.SheetFormatProperties.DefaultRowHeight.Value / 0.75 : double.NaN;

                        double[] columnWidths = new double[sheetDimension[2] - sheetDimension[0] + 1];
                        double[] rowHeights = new double[sheetDimension[3] - sheetDimension[1] + 1];
                        for (int columnWidthIndex = 0; columnWidthIndex < columnWidths.Length; columnWidthIndex++)
                        {
                            columnWidths[columnWidthIndex] = columnWidthDefault;
                        }
                        if (worksheet.GetFirstChild<Columns>() is Columns columnsGroup)
                        {
                            foreach (Column column in columnsGroup.Descendants<Column>())
                            {
                                bool isHidden = (column.Collapsed != null && column.Collapsed.HasValue && column.Collapsed.Value) || (column.Hidden != null && column.Hidden.HasValue && column.Hidden.Value);
                                if ((column.Width != null && column.Width.HasValue && (column.CustomWidth == null || (column.CustomWidth.HasValue && column.CustomWidth.Value))) || isHidden)
                                {
                                    for (int i = Math.Max(sheetDimension[0], column.Min != null && column.Min.HasValue ? (int)column.Min.Value : sheetDimension[0]); i <= Math.Min(sheetDimension[2], column.Max != null && column.Max.HasValue ? (int)column.Max.Value : sheetDimension[2]); i++)
                                    {
                                        columnWidths[i - sheetDimension[0]] = isHidden ? 0 : column.Width.Value;
                                    }
                                }
                            }
                        }

                        double columnWidthsTotal = columnWidths.Sum();
                        for (int columnWidthIndex = 0; columnWidthIndex < columnWidths.Length; columnWidthIndex++)
                        {
                            columnWidths[columnWidthIndex] = !double.IsNaN(columnWidthsTotal) ? columnWidths[columnWidthIndex] / columnWidthsTotal * 100 : columnWidths[columnWidthIndex] * 7;
                        }

                        int rowIndex = sheetDimension[1];
                        foreach (Row row in worksheet.Descendants<Row>())
                        {
                            rowIndex++;
                            if (row.RowIndex != null && row.RowIndex.HasValue)
                            {
                                if (row.RowIndex.Value < sheetDimension[1] || row.RowIndex.Value > sheetDimension[3])
                                {
                                    continue;
                                }
                                for (int additionalRowIndex = rowIndex; additionalRowIndex < row.RowIndex.Value; additionalRowIndex++)
                                {
                                    rowHeights[additionalRowIndex] = rowHeightDefault;
                                    writer.Write($"\n{new string(' ', 12)}<tr>");
                                    for (int additionalColumnIndex = 0; additionalColumnIndex < columnWidths.Length; additionalColumnIndex++)
                                    {
                                        writer.Write($"\n{new string(' ', 16)}<td style=\"height: {rowHeightDefault}px; width: {(!double.IsNaN(columnWidths[additionalColumnIndex]) ? $"{columnWidths[additionalColumnIndex]}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")}" : "auto")};\"></td>");
                                    }
                                    writer.Write($"\n{new string(' ', 12)}</tr>");
                                }
                                rowIndex = (int)row.RowIndex.Value;
                            }
                            double cellHeightActual = config.ConvertSizes ? ((row.CustomHeight == null || (row.CustomHeight.HasValue && row.CustomHeight.Value)) && row.Height != null && row.Height.HasValue ? row.Height.Value / 0.75 : rowHeightDefault) : double.NaN;
                            rowHeights[rowIndex - sheetDimension[1]] = cellHeightActual;

                            writer.Write($"\n{new string(' ', 12)}<tr>");

                            Cell[] cells = new Cell[columnWidths.Length];
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                if (cell.CellReference == null || !cell.CellReference.HasValue || GetRowIndex(cell.CellReference.Value) != rowIndex)
                                {
                                    continue;
                                }

                                int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
                                if (cellColumnIndex >= sheetDimension[0] && cellColumnIndex <= sheetDimension[2])
                                {
                                    cells[cellColumnIndex - sheetDimension[0]] = cell;
                                }
                            }
                            for (int additionalCellIndex = sheetDimension[0]; additionalCellIndex <= sheetDimension[2]; additionalCellIndex++)
                            {
                                if (cells[additionalCellIndex - sheetDimension[0]] != null)
                                {
                                    continue;
                                }

                                string additionalCellColumnName = string.Empty;
                                int additionalCellColumnIndex = additionalCellIndex;
                                while (additionalCellColumnIndex > 0)
                                {
                                    int modulo = (additionalCellColumnIndex - 1) % 26;
                                    additionalCellColumnName = (char)(65 + modulo) + additionalCellColumnName;
                                    additionalCellColumnIndex = (additionalCellColumnIndex - modulo) / 26;
                                }
                                cells[additionalCellIndex - sheetDimension[0]] = new Cell() { CellValue = new CellValue(string.Empty), CellReference = additionalCellColumnName + rowIndex };
                            }

                            int columnIndex = sheetDimension[0];
                            foreach (Cell cell in cells)
                            {
                                columnIndex = GetColumnIndex(cell.CellReference.Value);
                                double cellWidthActual = config.ConvertSizes ? columnWidths[columnIndex - sheetDimension[0]] : double.NaN;

                                int columnSpanned = 1;
                                int rowSpanned = 1;
                                if (isMergeCellsContained)
                                {
                                    if (mergeCells.Any(x => (x.FromColumn != columnIndex || x.FromRow != rowIndex) && columnIndex >= x.FromColumn && columnIndex <= x.ToColumn && rowIndex >= x.FromRow && rowIndex <= x.ToRow))
                                    {
                                        continue;
                                    }
                                    else if (mergeCells.FirstOrDefault(x => x.FromColumn == columnIndex && x.FromRow == rowIndex) is MergeCellInfo mergeCellInfo)
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
                                                int indexComponent = cellValueNumber > 0 || (numberFormatCodeComponents.Length == 2 && cellValueNumber == 0) ? 0 : (cellValueNumber < 0 ? 1 : (numberFormatCodeComponents.Length > 2 ? 2 : -1));
                                                numberFormatCode = indexComponent >= 0 ? numberFormatCodeComponents[indexComponent] : numberFormatCode;
                                            }
                                            else
                                            {
                                                numberFormatCode = numberFormatCodeComponents.Length > 3 ? numberFormatCodeComponents[3] : numberFormatCode;
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
                                if (config.ConvertStyles)
                                {
                                    if (!isNumberFormatDate)
                                    {
                                        if (cell.DataType != null && cell.DataType.HasValue)
                                        {
                                            if (cell.DataType.Value == CellValues.Error || cell.DataType.Value == CellValues.Boolean)
                                            {
                                                htmlStyleCell.Add("text-align", "center");
                                            }
                                            else if (cell.DataType.Value == CellValues.Number)
                                            {
                                                htmlStyleCell.Add("text-align", "right");
                                            }
                                        }
                                        else if (double.TryParse(cellValueRaw, out double _))
                                        {
                                            htmlStyleCell.Add("text-align", "right");
                                        }
                                    }
                                    if (styleIndex >= 0 && styleIndex < htmlStyleCellFormats.Length && htmlStyleCellFormats[styleIndex] != null)
                                    {
                                        htmlStyleCell = JoinHtmlAttributes(htmlStyleCell, htmlStyleCellFormats[styleIndex].Item1);
                                        cellValueContainer = cellValueContainer.Replace("{0}", htmlStyleCellFormats[styleIndex].Item2);
                                    }

                                    int differentialStyleIndex = -1;
                                    foreach (ConditionalFormatting conditionalFormatting in worksheet.Descendants<ConditionalFormatting>())
                                    {
                                        if (conditionalFormatting.SequenceOfReferences != null && conditionalFormatting.SequenceOfReferences.HasValue)
                                        {
                                            bool isFormattingApplicable = false;
                                            foreach (string references in conditionalFormatting.SequenceOfReferences.Items)
                                            {
                                                string[] range = references.Split(':');
                                                if (range.Length > 1)
                                                {
                                                    int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
                                                    int cellRowIndex = GetRowIndex(cell.CellReference.Value);
                                                    GetReferenceRange(range, out int referenceFromColumn, out int referenceFromRow, out int referenceToColumn, out int referenceToRow);
                                                    if (cellColumnIndex >= referenceFromColumn && cellColumnIndex <= referenceToColumn && cellRowIndex >= referenceFromRow && cellRowIndex <= referenceToRow)
                                                    {
                                                        isFormattingApplicable = true;
                                                        break;
                                                    }
                                                }
                                                else if (cell.CellReference != null && cell.CellReference.HasValue && cell.CellReference.Value == references)
                                                {
                                                    isFormattingApplicable = true;
                                                    break;
                                                }
                                            }
                                            if (!isFormattingApplicable)
                                            {
                                                continue;
                                            }
                                        }

                                        int priorityCurrent = int.MaxValue;
                                        foreach (ConditionalFormattingRule formattingRule in conditionalFormatting.Descendants<ConditionalFormattingRule>())
                                        {
                                            if (formattingRule.FormatId == null || !formattingRule.FormatId.HasValue || formattingRule.Type == null || !formattingRule.Type.HasValue)
                                            {
                                                continue;
                                            }
                                            else if (formattingRule.Priority != null && formattingRule.Priority.HasValue)
                                            {
                                                if (formattingRule.Priority.Value > priorityCurrent)
                                                {
                                                    continue;
                                                }
                                                priorityCurrent = formattingRule.Priority.Value;
                                            }

                                            //TODO: complete rules
                                            bool isConditionMet = false;
                                            if (formattingRule.Type.Value == ConditionalFormatValues.CellIs && formattingRule.GetFirstChild<Formula>() is Formula formula)
                                            {
                                                string formulaValue = formula.Text.Trim('"');
                                                switch (formattingRule.Operator != null && formattingRule.Operator.HasValue ? formattingRule.Operator.Value : ConditionalFormattingOperatorValues.Equal)
                                                {
                                                    case ConditionalFormattingOperatorValues.Equal:
                                                        isConditionMet = cellValueRaw == formulaValue;
                                                        break;
                                                    case ConditionalFormattingOperatorValues.NotEqual:
                                                        isConditionMet = cellValueRaw != formulaValue;
                                                        break;
                                                    case ConditionalFormattingOperatorValues.BeginsWith:
                                                        isConditionMet = cellValueRaw.StartsWith(formulaValue);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.EndsWith:
                                                        isConditionMet = cellValueRaw.EndsWith(formulaValue);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.ContainsText:
                                                        isConditionMet = cellValueRaw.Contains(formulaValue);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.NotContains:
                                                        isConditionMet = !cellValueRaw.Contains(formulaValue);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.GreaterThan:
                                                        isConditionMet = double.TryParse(formulaValue, out double formulaValueNumberGreaterThan) && double.TryParse(cellValueRaw, out double cellValueNumberGreaterThan) && cellValueNumberGreaterThan > formulaValueNumberGreaterThan;
                                                        break;
                                                    case ConditionalFormattingOperatorValues.GreaterThanOrEqual:
                                                        isConditionMet = double.TryParse(formulaValue, out double formulaValueNumberGreaterThanOrEqual) && double.TryParse(cellValueRaw, out double cellValueNumberGreaterThanOrEqual) && cellValueNumberGreaterThanOrEqual >= formulaValueNumberGreaterThanOrEqual;
                                                        break;
                                                    case ConditionalFormattingOperatorValues.LessThan:
                                                        isConditionMet = double.TryParse(formulaValue, out double formulaValueNumberLessThan) && double.TryParse(cellValueRaw, out double cellValueNumberLessThan) && cellValueNumberLessThan < formulaValueNumberLessThan;
                                                        break;
                                                    case ConditionalFormattingOperatorValues.LessThanOrEqual:
                                                        isConditionMet = double.TryParse(formulaValue, out double formulaValueNumberLessThanOrEqual) && double.TryParse(cellValueRaw, out double cellValueNumberLessThanOrEqual) && cellValueNumberLessThanOrEqual <= formulaValueNumberLessThanOrEqual;
                                                        break;
                                                }
                                            }
                                            else if (formattingRule.Text != null && formattingRule.Text.HasValue)
                                            {
                                                switch (formattingRule.Type.Value)
                                                {
                                                    case ConditionalFormatValues.BeginsWith:
                                                        isConditionMet = cellValueRaw.StartsWith(formattingRule.Text.Value);
                                                        break;
                                                    case ConditionalFormatValues.EndsWith:
                                                        isConditionMet = cellValueRaw.EndsWith(formattingRule.Text.Value);
                                                        break;
                                                    case ConditionalFormatValues.ContainsText:
                                                        isConditionMet = cellValueRaw.Contains(formattingRule.Text.Value);
                                                        break;
                                                    case ConditionalFormatValues.NotContainsText:
                                                        isConditionMet = !cellValueRaw.Contains(formattingRule.Text.Value);
                                                        break;
                                                }
                                            }

                                            if (isConditionMet)
                                            {
                                                differentialStyleIndex = (int)formattingRule.FormatId.Value;
                                            }
                                        }
                                    }
                                    if (differentialStyleIndex >= 0 && differentialStyleIndex < htmlStyleDifferentialFormats.Length && htmlStyleDifferentialFormats[differentialStyleIndex] != null)
                                    {
                                        htmlStyleCell = JoinHtmlAttributes(htmlStyleCell, htmlStyleDifferentialFormats[differentialStyleIndex].Item1);
                                        cellValueContainer = cellValueContainer.Replace("{0}", htmlStyleDifferentialFormats[differentialStyleIndex].Item2);
                                    }
                                }

                                writer.Write($"\n{new string(' ', 16)}<td{(columnSpanned != 1 ? $" colspan=\"{columnSpanned}\"" : string.Empty)}{(rowSpanned != 1 ? $" rowspan=\"{rowSpanned}\"" : string.Empty)} style=\"width: {(!double.IsNaN(cellWidthActual) ? $"{cellWidthActual}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")}" : "auto")}; height: {(double.IsNaN(cellHeightActual) ? "auto" : $"{cellHeightActual}px")};{GetHtmlAttributesString(htmlStyleCell, true)}\">{cellValueContainer.Replace("{0}", cellValue)}</td>");
                            }

                            writer.Write($"\n{new string(' ', 12)}</tr>");

                            progressCallback?.Invoke(document, new ConverterProgressCallbackEventArgs(sheetIndex, sheetsCount, rowIndex, rowHeights.Length));
                        }

                        writer.Write($"\n{new string(' ', 8)}</table>");

                        if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                        {
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absoluteAnchor in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor>())
                            {
                                string left = absoluteAnchor.Position != null && absoluteAnchor.Position.X != null && absoluteAnchor.Position.X.HasValue ? $"{(double)absoluteAnchor.Position.X.Value / 914400 * 96}px" : "auto";
                                string top = absoluteAnchor.Position != null && absoluteAnchor.Position.Y != null && absoluteAnchor.Position.Y.HasValue ? $"{(double)absoluteAnchor.Position.Y.Value / 914400 * 96}px" : "auto";
                                string width = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cx != null && absoluteAnchor.Extent.Cx.HasValue ? $"{(double)absoluteAnchor.Extent.Cx.Value / 914400 * 96}px" : "auto";
                                string height = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cy != null && absoluteAnchor.Extent.Cy.HasValue ? $"{(double)absoluteAnchor.Extent.Cy.Value / 914400 * 96}px" : "auto";
                                DrawingsToHtml(worksheetPart, absoluteAnchor, writer, left, top, width, height, config.ConvertPictures);
                            }
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>())
                            {
                                double left = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) ? columnWidths.Take(Math.Min(columnWidths.Length, columnId - sheetDimension[0] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double leftOffset = oneCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(oneCellAnchor.FromMarker.ColumnOffset.Text, out double columnOffset) ? columnOffset / 914400 * 96 : 0;
                                double top = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) ? rowHeights.Take(Math.Min(rowHeights.Length, rowId - sheetDimension[1] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double topOffset = oneCellAnchor.FromMarker.RowOffset != null && double.TryParse(oneCellAnchor.FromMarker.RowOffset.Text, out double rowOffset) ? rowOffset / 914400 * 96 : 0;
                                string width = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cx != null && oneCellAnchor.Extent.Cx.HasValue ? $"{(double)oneCellAnchor.Extent.Cx.Value / 914400 * 96}px" : "auto";
                                string height = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cy != null && oneCellAnchor.Extent.Cy.HasValue ? $"{(double)oneCellAnchor.Extent.Cy.Value / 914400 * 96}px" : "auto";
                                DrawingsToHtml(worksheetPart, oneCellAnchor, writer, $"{(!double.IsNaN(left) ? $"calc({left}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")} + {leftOffset}px)" : $"{leftOffset}px")}", $"{(!double.IsNaN(top) ? top + topOffset : topOffset)}px", width, height, config.ConvertPictures);
                            }
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                            {
                                double fromColumn = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) ? columnWidths.Take(Math.Min(columnWidths.Length, fromColumnId - sheetDimension[0] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double fromColumnOffset = twoCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.FromMarker.ColumnOffset.Text, out double fromMarkerColumnOffset) ? fromMarkerColumnOffset / 914400 * 96 : 0;
                                double fromRow = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) ? rowHeights.Take(Math.Min(rowHeights.Length, fromRowId - sheetDimension[1] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double fromRowOffset = twoCellAnchor.FromMarker.RowOffset != null && double.TryParse(twoCellAnchor.FromMarker.RowOffset.Text, out double fromMarkerRowOffset) ? fromMarkerRowOffset / 914400 * 96 : 0;
                                double toColumn = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) ? columnWidths.Take(Math.Min(columnWidths.Length, toColumnId - sheetDimension[0] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double toColumnOffset = twoCellAnchor.ToMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.ToMarker.ColumnOffset.Text, out double toMarkerColumnOffset) ? toMarkerColumnOffset / 914400 * 96 : 0;
                                double toRow = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) ? rowHeights.Take(Math.Min(rowHeights.Length, toRowId - sheetDimension[1] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double toRowOffset = twoCellAnchor.ToMarker.RowOffset != null && double.TryParse(twoCellAnchor.ToMarker.RowOffset.Text, out double toMarkerRowOffset) ? toMarkerRowOffset / 914400 * 96 : 0;
                                string leftCalculation = $"{(!double.IsNaN(fromColumn) ? $"calc({fromColumn}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")} + {fromColumnOffset}px)" : $"{fromColumnOffset}px")}";
                                string topCalculation = $"{(!double.IsNaN(fromRow) ? fromRow + fromRowOffset : fromRowOffset)}px";
                                DrawingsToHtml(worksheetPart, twoCellAnchor, writer, leftCalculation, topCalculation, $"calc({(!double.IsNaN(toColumn) ? $"calc({toColumn}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")} + {toColumnOffset}px)" : $"{toColumnOffset}px")} - {leftCalculation})", $"calc({$"{(!double.IsNaN(toRow) ? toRow + toRowOffset : toRowOffset)}px"} - {topCalculation})", config.ConvertPictures);
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
            return Math.Max(0, columnNumber + 1);
        }
        private static int GetRowIndex(string cellName)
        {
            Match match = regexNumbers.Match(cellName);
            return match.Success && int.TryParse(match.Value, out int rowIndex) ? rowIndex : 0;
        }

        private static void GetReferenceRange(string[] range, out int fromColumn, out int fromRow, out int toColumn, out int toRow)
        {
            int firstColumn = range.Length > 1 ? GetColumnIndex(range[0]) : 0;
            int firstRow = range.Length > 1 ? GetRowIndex(range[0]) : 0;
            int secondColumn = range.Length > 1 ? GetColumnIndex(range[1]) : 0;
            int secondRow = range.Length > 1 ? GetRowIndex(range[1]) : 0;
            fromColumn = Math.Min(firstColumn, secondColumn);
            fromRow = Math.Min(firstRow, secondRow);
            toColumn = Math.Max(firstColumn, secondColumn);
            toRow = Math.Max(firstRow, secondRow);
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
            return isAdditional ? $" {htmlAttributes.Trim()}" : htmlAttributes.Trim();
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

            object[] infoScientific = null;
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
                        isPeriodRequired = (indexes[5] > indexes[2] && (formatChar == '0' || formatChar == '?') && infoScientific == null) || isPeriodRequired;
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
                            infoScientific = new object[] { true, (indexes[0] - 1).ToString() };
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
                                infoScientific = new object[] { false, (digit - indexes[0]).ToString() };
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
                    else if (isValueNumber && formatChar == 'E' && isIncreasing && indexes[5] + 1 < format.Length && (format[indexes[5] + 1] == '+' || format[indexes[5] + 1] == '-') && infoScientific != null && infoScientific.Length > 1)
                    {
                        resultFormatted = result + (infoScientific[0] is bool isPositive && isPositive ? (format[indexes[5] + 1] == '-' ? "E" : "E+") : "E-");
                        result = string.Empty;
                        isIncreasing = false;
                        indexes[4] = (infoScientific[1] as string ?? "").Length;
                        indexes[5] = indexes[3];
                        isFormattingScientific = true;
                        continue;
                    }
                    else if (isValueNumber && (formatChar == '0' || formatChar == '#' || formatChar == '?'))
                    {
                        indexes[4] += isIncreasing ? 1 : -1;
                        if (indexes[4] >= 0 && indexes[4] < (!isFormattingScientific ? value.Length : (infoScientific[1] as string ?? "").Length) && (formatChar == '0' || indexes[4] > 0 || value[indexes[4]] != '0' || isPeriodRequired || isFormattingScientific))
                        {
                            if (isIncreasing && (indexes[5] >= indexes[3] || (indexes[5] + 2 < format.Length && format[indexes[5] + 1] == 'E' && (format[indexes[5] + 2] == '+' || format[indexes[5] + 2] == '-'))) && indexes[4] + 1 < value.Length && int.TryParse(value[indexes[4] + 1].ToString(), out int next) && next > 4)
                            {
                                return GetFormattedNumber((valueNumber + (10 - next) / Math.Pow(10, indexes[4] + 1 - indexes[0])).ToString(), format);
                            }

                            result += !isFormattingScientific ? value[indexes[4]].ToString() : (infoScientific[1] as string ?? "")[indexes[4]].ToString();
                            if (!isFormattingScientific ? indexes[5] <= indexes[1] : (indexes[5] - 2 >= 0 && format[indexes[5] - 2] == 'E' && (format[indexes[5] - 1] == '+' || format[indexes[5] - 1] == '-')))
                            {
                                result += new string((!isFormattingScientific ? value : (infoScientific[1] as string ?? "")).Substring(0, indexes[4]).Reverse().ToArray());
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

            if (fill != null && fill.PatternFill != null && (fill.PatternFill.PatternType == null || (fill.PatternFill.PatternType.HasValue && fill.PatternFill.PatternType.Value != PatternValues.None)))
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
                if (alignment.Horizontal != null && alignment.Horizontal.HasValue && alignment.Horizontal.Value != HorizontalAlignmentValues.General)
                {
                    htmlStyle.Add("text-align", alignment.Horizontal.Value == HorizontalAlignmentValues.Left ? "left" : (alignment.Horizontal.Value == HorizontalAlignmentValues.Right ? "right" : (alignment.Horizontal.Value == HorizontalAlignmentValues.Justify ? "justify" : "center")));
                }
                if (alignment.Vertical != null && alignment.Vertical.HasValue)
                {
                    htmlStyle.Add("vertical-align", alignment.Vertical.Value == VerticalAlignmentValues.Bottom ? "bottom" : (alignment.Vertical.Value == VerticalAlignmentValues.Top ? "top" : "middle"));
                }
                if (alignment.WrapText != null && alignment.WrapText.HasValue && alignment.WrapText.Value)
                {
                    htmlStyle.Add("word-wrap", "break-word");
                    htmlStyle.Add("white-space", "normal");
                }
                if (alignment.TextRotation != null && alignment.TextRotation.HasValue)
                {
                    cellValueContainer = cellValueContainer.Replace("{0}", $"<div style=\"width: fit-content; transform: rotate(-{alignment.TextRotation.Value}deg);\">{{0}}</div>");
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
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                        break;
                    case 1:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
                        break;
                    case 2:
                        rgbColor = new RgbaColor() { R = 255, G = 0, B = 0, A = 1 };
                        break;
                    case 3:
                        rgbColor = new RgbaColor() { R = 0, G = 255, B = 0, A = 1 };
                        break;
                    case 4:
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 255, A = 1 };
                        break;
                    case 5:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 0, A = 1 };
                        break;
                    case 6:
                        rgbColor = new RgbaColor() { R = 255, G = 0, B = 255, A = 1 };
                        break;
                    case 7:
                        rgbColor = new RgbaColor() { R = 0, G = 255, B = 255, A = 1 };
                        break;
                    case 8:
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };
                        break;
                    case 9:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
                        break;
                    case 10:
                        rgbColor = new RgbaColor() { R = 255, G = 0, B = 0, A = 1 };
                        break;
                    case 11:
                        rgbColor = new RgbaColor() { R = 0, G = 255, B = 0, A = 1 };
                        break;
                    case 12:
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 255, A = 1 };
                        break;
                    case 13:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 0, A = 1 };
                        break;
                    case 14:
                        rgbColor = new RgbaColor() { R = 255, G = 0, B = 255, A = 1 };
                        break;
                    case 15:
                        rgbColor = new RgbaColor() { R = 0, G = 255, B = 255, A = 1 };
                        break;
                    case 16:
                        rgbColor = new RgbaColor() { R = 128, G = 0, B = 0, A = 1 };
                        break;
                    case 17:
                        rgbColor = new RgbaColor() { R = 0, G = 128, B = 0, A = 1 };
                        break;
                    case 18:
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 128, A = 1 };
                        break;
                    case 19:
                        rgbColor = new RgbaColor() { R = 128, G = 128, B = 0, A = 1 };
                        break;
                    case 20:
                        rgbColor = new RgbaColor() { R = 128, G = 0, B = 128, A = 1 };
                        break;
                    case 21:
                        rgbColor = new RgbaColor() { R = 0, G = 128, B = 128, A = 1 };
                        break;
                    case 22:
                        rgbColor = new RgbaColor() { R = 192, G = 192, B = 192, A = 1 };
                        break;
                    case 23:
                        rgbColor = new RgbaColor() { R = 128, G = 128, B = 128, A = 1 };
                        break;
                    case 24:
                        rgbColor = new RgbaColor() { R = 153, G = 153, B = 255, A = 1 };
                        break;
                    case 25:
                        rgbColor = new RgbaColor() { R = 153, G = 51, B = 102, A = 1 };
                        break;
                    case 26:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 204, A = 1 };
                        break;
                    case 27:
                        rgbColor = new RgbaColor() { R = 204, G = 255, B = 255, A = 1 };
                        break;
                    case 28:
                        rgbColor = new RgbaColor() { R = 102, G = 0, B = 102, A = 1 };
                        break;
                    case 29:
                        rgbColor = new RgbaColor() { R = 255, G = 128, B = 128, A = 1 };
                        break;
                    case 30:
                        rgbColor = new RgbaColor() { R = 0, G = 102, B = 204, A = 1 };
                        break;
                    case 31:
                        rgbColor = new RgbaColor() { R = 204, G = 204, B = 255, A = 1 };
                        break;
                    case 32:
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 128, A = 1 };
                        break;
                    case 33:
                        rgbColor = new RgbaColor() { R = 255, G = 0, B = 255, A = 1 };
                        break;
                    case 34:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 0, A = 1 };
                        break;
                    case 35:
                        rgbColor = new RgbaColor() { R = 0, G = 255, B = 255, A = 1 };
                        break;
                    case 36:
                        rgbColor = new RgbaColor() { R = 128, G = 0, B = 128, A = 1 };
                        break;
                    case 37:
                        rgbColor = new RgbaColor() { R = 128, G = 0, B = 0, A = 1 };
                        break;
                    case 38:
                        rgbColor = new RgbaColor() { R = 0, G = 128, B = 128, A = 1 };
                        break;
                    case 39:
                        rgbColor = new RgbaColor() { R = 0, G = 0, B = 255, A = 1 };
                        break;
                    case 40:
                        rgbColor = new RgbaColor() { R = 0, G = 204, B = 255, A = 1 };
                        break;
                    case 41:
                        rgbColor = new RgbaColor() { R = 204, G = 255, B = 255, A = 1 };
                        break;
                    case 42:
                        rgbColor = new RgbaColor() { R = 204, G = 255, B = 204, A = 1 };
                        break;
                    case 43:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 153, A = 1 };
                        break;
                    case 44:
                        rgbColor = new RgbaColor() { R = 153, G = 204, B = 255, A = 1 };
                        break;
                    case 45:
                        rgbColor = new RgbaColor() { R = 255, G = 153, B = 204, A = 1 };
                        break;
                    case 46:
                        rgbColor = new RgbaColor() { R = 204, G = 153, B = 255, A = 1 };
                        break;
                    case 47:
                        rgbColor = new RgbaColor() { R = 255, G = 204, B = 153, A = 1 };
                        break;
                    case 48:
                        rgbColor = new RgbaColor() { R = 51, G = 102, B = 255, A = 1 };
                        break;
                    case 49:
                        rgbColor = new RgbaColor() { R = 51, G = 204, B = 204, A = 1 };
                        break;
                    case 50:
                        rgbColor = new RgbaColor() { R = 153, G = 204, B = 0, A = 1 };
                        break;
                    case 51:
                        rgbColor = new RgbaColor() { R = 255, G = 204, B = 0, A = 1 };
                        break;
                    case 52:
                        rgbColor = new RgbaColor() { R = 255, G = 153, B = 0, A = 1 };
                        break;
                    case 53:
                        rgbColor = new RgbaColor() { R = 255, G = 102, B = 0, A = 1 };
                        break;
                    case 54:
                        rgbColor = new RgbaColor() { R = 102, G = 102, B = 153, A = 1 };
                        break;
                    case 55:
                        rgbColor = new RgbaColor() { R = 150, G = 150, B = 150, A = 1 };
                        break;
                    case 56:
                        rgbColor = new RgbaColor() { R = 0, G = 51, B = 102, A = 1 };
                        break;
                    case 57:
                        rgbColor = new RgbaColor() { R = 51, G = 153, B = 102, A = 1 };
                        break;
                    case 58:
                        rgbColor = new RgbaColor() { R = 0, G = 51, B = 0, A = 1 };
                        break;
                    case 59:
                        rgbColor = new RgbaColor() { R = 51, G = 51, B = 0, A = 1 };
                        break;
                    case 60:
                        rgbColor = new RgbaColor() { R = 153, G = 51, B = 0, A = 1 };
                        break;
                    case 61:
                        rgbColor = new RgbaColor() { R = 153, G = 51, B = 102, A = 1 };
                        break;
                    case 62:
                        rgbColor = new RgbaColor() { R = 51, G = 51, B = 153, A = 1 };
                        break;
                    case 63:
                        rgbColor = new RgbaColor() { R = 51, G = 51, B = 51, A = 1 };
                        break;
                    case 64:
                        rgbColor = new RgbaColor() { R = 128, G = 128, B = 128, A = 1 };
                        break;
                    case 65:
                        rgbColor = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
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
                    HlsToRgb(color.HslColor.HueValue.HasValue ? color.HslColor.HueValue.Value / 6000.0 : 0, color.HslColor.LumValue.HasValue ? color.HslColor.LumValue.Value : 0, color.HslColor.SatValue.HasValue ? color.HslColor.SatValue.Value : 0, out double red, out double green, out double blue);
                    rgbColor.R = (int)red;
                    rgbColor.G = (int)green;
                    rgbColor.B = (int)blue;
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

            if (type.Tint != null && type.Tint.HasValue && type.Tint.Value != 0)
            {
                RgbToHls(rgbColor.R, rgbColor.G, rgbColor.B, out double hue, out double luminosity, out double saturation);
                luminosity = type.Tint.Value < 0 ? luminosity * (1 + type.Tint.Value) : luminosity * (1 - type.Tint.Value) + type.Tint.Value;
                HlsToRgb(hue, luminosity, saturation, out double red, out double green, out double blue);
                rgbColor.R = (int)red;
                rgbColor.G = (int)green;
                rgbColor.B = (int)blue;
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
            return new RgbaColor()
            {
                R = hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 2 : 0, 2), 16) : 0,
                G = hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 4 : 2, 2), 16) : 0,
                B = hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 6 : 4, 2), 16) : 0,
                A = hexTrimmed.Length > 5 ? (hexTrimmed.Length > 7 ? Convert.ToInt32(hexTrimmed.Substring(0, 2), 16) / 255.0 : 1) : 0
            };
        }

        private static void RgbToHls(double red, double green, double blue, out double hue, out double luminosity, out double saturation)
        {
            double redMapped = red / 255;
            double greenMapped = green / 255;
            double blueMapped = blue / 255;

            double max = Math.Max(redMapped, Math.Max(greenMapped, blueMapped));
            double min = Math.Min(redMapped, Math.Min(greenMapped, blueMapped));
            double chroma = max - min;
            double distanceRed = (max - redMapped) / chroma;
            double distanceGreen = (max - greenMapped) / chroma;
            double distanceBlue = (max - blueMapped) / chroma;
            hue = chroma == 0 ? 0 : ((redMapped == max ? distanceBlue - distanceGreen : (greenMapped == max ? 2 + distanceRed - distanceBlue : 4 + distanceGreen - distanceRed)) * 60 % 360 + 360) % 360;
            luminosity = (max + min) / 2;
            saturation = chroma == 0 ? 0 : (luminosity <= 0.5 ? chroma / (max + min) : chroma / (2 - max - min));
        }

        private static void HlsToRgb(double hue, double luminosity, double saturation, out double red, out double green, out double blue)
        {
            double value1 = luminosity <= 0.5 ? luminosity * (1 + saturation) : luminosity + saturation - luminosity * saturation;
            double value2 = 2 * luminosity - value1;
            Func<double, double> actionCalculateColor = (hueShifted) =>
            {
                hueShifted = (hueShifted % 360 + 360) % 360;
                return hueShifted < 60 ? value2 + (value1 - value2) * hueShifted / 60 : (hueShifted < 180 ? value1 : (hueShifted < 240 ? value2 + (value1 - value2) * (240 - hueShifted) / 60 : value2));
            };
            red = (saturation == 0 ? luminosity : actionCalculateColor.Invoke(hue + 120)) * 255.0;
            green = (saturation == 0 ? luminosity : actionCalculateColor.Invoke(hue)) * 255.0;
            blue = (saturation == 0 ? luminosity : actionCalculateColor.Invoke(hue - 120)) * 255.0;
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
                htmlStyle.Add("font-size", $"{fontSize.Val.Value / 72 * 96}px");
            }
            if (bold != null)
            {
                htmlStyle.Add("font-weight", bold.Val == null || (bold.Val.HasValue && bold.Val.Value) ? "bold" : "normal");
            }
            if (italic != null)
            {
                htmlStyle.Add("font-style", italic.Val == null || (italic.Val.HasValue && italic.Val.Value) ? "italic" : "normal");
            }
            string htmlStyleTextDecoraion = string.Empty;
            if (strike != null)
            {
                htmlStyleTextDecoraion += strike.Val == null || (strike.Val.HasValue && strike.Val.Value) ? " line-through" : " none";
            }
            if (underline != null && underline.Val != null && underline.Val.HasValue)
            {
                if (underline.Val.Value == UnderlineValues.Double || underline.Val.Value == UnderlineValues.DoubleAccounting)
                {
                    htmlStyleTextDecoraion += " underline double";
                }
                else if (underline.Val.Value != UnderlineValues.None)
                {
                    htmlStyleTextDecoraion += " underline";
                }
            }
            if (!string.IsNullOrEmpty(htmlStyleTextDecoraion))
            {
                htmlStyle.Add("text-decoration", htmlStyleTextDecoraion.Trim());
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
                    case BorderStyleValues.Thin:
                        width = "thin";
                        style = "solid";
                        break;
                    case BorderStyleValues.Thick:
                        width = "thick";
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
                        style = "dashed";
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

        private static void DrawingsToHtml(WorksheetPart worksheet, OpenXmlCompositeElement anchor, StreamWriter writer, string left, string top, string width, string height, bool isPictureAllowed)
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
                            Prefix = $"<img src=\"data:{imagePart.ContentType};base64,{base64}\"{(properties != null && properties.Description != null && properties.Description.HasValue ? $" alt=\"{properties.Description.Value}\"" : string.Empty)}",
                            Postfix = "/>",
                            Left = left,
                            Top = top,
                            Width = width,
                            Height = height,
                            IsHidden = properties != null && properties.Hidden != null && properties.Hidden.HasValue && properties.Hidden.Value,
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
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height,
                    IsHidden = properties != null && properties.Hidden != null && properties.Hidden.HasValue && properties.Hidden.Value,
                    ShapeProperties = shape.ShapeProperties
                });
            }

            foreach (DrawingInfo drawingInfo in drawings)
            {
                string widthActual = drawingInfo.Width;
                string heightActual = drawingInfo.Height;
                string htmlStyleTransform = string.Empty;
                if (drawingInfo.ShapeProperties != null && drawingInfo.ShapeProperties.Transform2D != null)
                {
                    if (drawingInfo.ShapeProperties.Transform2D.Extents != null)
                    {
                        //TODO: fix
                        if (false)
                        {
                            if (drawingInfo.ShapeProperties.Transform2D.Extents.Cx != null && drawingInfo.ShapeProperties.Transform2D.Extents.Cx.HasValue)
                            {
                                widthActual = $"{drawingInfo.ShapeProperties.Transform2D.Extents.Cx.Value / 914400 * 96}px";
                            }
                            if (drawingInfo.ShapeProperties.Transform2D.Extents.Cy != null && drawingInfo.ShapeProperties.Transform2D.Extents.Cy.HasValue)
                            {
                                heightActual = $"{drawingInfo.ShapeProperties.Transform2D.Extents.Cy.Value / 914400 * 96}px";
                            }
                        }
                    }
                    if (drawingInfo.ShapeProperties.Transform2D.Offset != null)
                    {
                        if (drawingInfo.ShapeProperties.Transform2D.Offset.X != null && drawingInfo.ShapeProperties.Transform2D.Offset.X.HasValue)
                        {
                            htmlStyleTransform += $" translateX({drawingInfo.ShapeProperties.Transform2D.Offset.X.Value / 914400 * 96}px)";
                        }
                        if (drawingInfo.ShapeProperties.Transform2D.Offset.Y != null && drawingInfo.ShapeProperties.Transform2D.Offset.Y.HasValue)
                        {
                            htmlStyleTransform += $" translateY({drawingInfo.ShapeProperties.Transform2D.Offset.Y.Value / 914400 * 96}px)";
                        }
                    }
                    if (drawingInfo.ShapeProperties.Transform2D.Rotation != null && drawingInfo.ShapeProperties.Transform2D.Rotation.HasValue)
                    {
                        htmlStyleTransform += $" rotate(-{drawingInfo.ShapeProperties.Transform2D.Rotation.Value}deg)";
                    }
                }

                writer.Write($"\n{new string(' ', 8)}{drawingInfo.Prefix} style=\"position: absolute; left: {drawingInfo.Left}; top: {drawingInfo.Top}; width: {widthActual}; height: {heightActual};{(!string.IsNullOrEmpty(htmlStyleTransform) ? $" transform:{htmlStyleTransform};" : string.Empty)}{(drawingInfo.IsHidden ? " visibility: hidden;" : string.Empty)}\"{drawingInfo.Postfix}");
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
            public string Left { get; set; }
            public string Top { get; set; }
            public string Width { get; set; }
            public string Height { get; set; }
            public bool IsHidden { get; set; }
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
            this.ConvertStyles = true;
            this.ConvertSizes = true;
            this.ConvertPictures = true;
            this.ConvertSheetTitles = true;
            this.ConvertHiddenSheets = false;
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
        public bool ConvertStyles { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx cell sizes into Html table cell sizes or not.
        /// </summary>
        public bool ConvertSizes { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx pictures into Html pictures or not.
        /// </summary>
        public bool ConvertPictures { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx sheet names to titles or not.
        /// </summary>
        public bool ConvertSheetTitles { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx hidden sheets or not.
        /// </summary>
        public bool ConvertHiddenSheets { get; set; }

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
