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
    /// The Xlsx to Html converter.
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

            StreamWriter writer = new StreamWriter(outputHtml, config.Encoding, config.BufferSize);
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
                            htmlStyleCellFormats[stylesheetFormatIndex] = new Tuple<Dictionary<string, string>, string>(CellFormatToHtml(workbook, fill, font, border, cellFormat.ApplyAlignment == null || (cellFormat.ApplyAlignment.HasValue && cellFormat.ApplyAlignment.Value) ? cellFormat.Alignment : null, out string cellValueContainer, config), cellValueContainer);
                        }
                    }
                    Tuple<Dictionary<string, string>, string>[] htmlStyleDifferentialFormats = new Tuple<Dictionary<string, string>, string>[stylesheet != null && stylesheet.DifferentialFormats != null && stylesheet.DifferentialFormats.HasChildren ? stylesheet.DifferentialFormats.ChildElements.Count : 0];
                    for (int stylesheetDifferentialFormatIndex = 0; stylesheetDifferentialFormatIndex < htmlStyleDifferentialFormats.Length; stylesheetDifferentialFormatIndex++)
                    {
                        if (stylesheet.DifferentialFormats.ChildElements[stylesheetDifferentialFormatIndex] is DifferentialFormat differentialFormat)
                        {
                            htmlStyleDifferentialFormats[stylesheetDifferentialFormatIndex] = new Tuple<Dictionary<string, string>, string>(CellFormatToHtml(workbook, differentialFormat.Fill, differentialFormat.Font, differentialFormat.Border, differentialFormat.Alignment, out string cellValueContainer, config), cellValueContainer);
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
                                        string cellValueContainer = "{0}";
                                        if (config.ConvertStyles && run.RunProperties is RunProperties runProperties)
                                        {
                                            if (runProperties.GetFirstChild<RunFont>() is RunFont runFont && runFont.Val != null && runFont.Val.HasValue)
                                            {
                                                htmlStyleRun.Add("font-family", runFont.Val.Value);
                                            }
                                            htmlStyleRun = JoinHtmlAttributes(htmlStyleRun, FontToHtml(workbook, runProperties.GetFirstChild<Color>(), runProperties.GetFirstChild<FontSize>(), runProperties.GetFirstChild<Bold>(), runProperties.GetFirstChild<Italic>(), runProperties.GetFirstChild<Strike>(), runProperties.GetFirstChild<Underline>(), out cellValueContainer, config));
                                        }

                                        string runContent = string.IsNullOrEmpty(run.Text.Text) ? run.Text.InnerText : run.Text.Text;
                                        cellValue += $"<p style=\"display: inline;{GetHtmlAttributesString(htmlStyleRun, true)}\">{cellValueContainer.Replace("{0}", GetEscapedString(runContent))}</p>";
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
                            string tabColor = worksheet.SheetProperties != null && worksheet.SheetProperties.TabColor != null ? ColorTypeToHtml(workbook, worksheet.SheetProperties.TabColor, config) : string.Empty;
                            writer.Write($"\n{new string(' ', 4)}<h5{(!string.IsNullOrEmpty(tabColor) ? $" style=\"border-bottom-color: {tabColor};\"" : string.Empty)}>{(sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : "Untitled")}</h5>");
                        }

                        writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                        writer.Write($"\n{new string(' ', 8)}<table>");

                        int[] sheetDimension = new int[4] { 1, 1, 1, 1 };
                        if (worksheet.SheetDimension != null && worksheet.SheetDimension.Reference != null && worksheet.SheetDimension.Reference.HasValue)
                        {
                            GetReferenceRange(worksheet.SheetDimension.Reference.Value, out sheetDimension[0], out sheetDimension[1], out sheetDimension[2], out sheetDimension[3]);
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

                        List<int[]> mergeCells = new List<int[]>();
                        if (worksheet.Descendants<MergeCells>().FirstOrDefault() is MergeCells mergeCellsGroup)
                        {
                            foreach (MergeCell mergeCell in mergeCellsGroup.Descendants<MergeCell>())
                            {
                                if (mergeCell.Reference != null && mergeCell.Reference.HasValue)
                                {
                                    GetReferenceRange(mergeCell.Reference.Value, out int mergeCellFromColumn, out int mergeCellFromRow, out int mergeCellToColumn, out int mergeCellToRow);
                                    mergeCells.Add(new int[6] { mergeCellFromColumn, mergeCellFromRow, mergeCellToColumn, mergeCellToRow, mergeCellToColumn - mergeCellFromColumn + 1, mergeCellToRow - mergeCellFromRow + 1 });
                                }
                            }
                        }

                        IEnumerable<ConditionalFormatting> conditionalFormattings = worksheet.Descendants<ConditionalFormatting>();

                        double columnWidthDefault = worksheet.SheetFormatProperties != null && worksheet.SheetFormatProperties.DefaultColumnWidth != null && worksheet.SheetFormatProperties.DefaultColumnWidth.HasValue ? worksheet.SheetFormatProperties.DefaultColumnWidth.Value : (worksheet.SheetFormatProperties != null && worksheet.SheetFormatProperties.BaseColumnWidth != null && worksheet.SheetFormatProperties.BaseColumnWidth.HasValue ? worksheet.SheetFormatProperties.BaseColumnWidth.Value : double.NaN);
                        double rowHeightDefault = worksheet.SheetFormatProperties != null && worksheet.SheetFormatProperties.DefaultRowHeight != null && worksheet.SheetFormatProperties.DefaultRowHeight.HasValue ? worksheet.SheetFormatProperties.DefaultRowHeight.Value / 0.75 : double.NaN;

                        double[] columnWidths = new double[sheetDimension[2] - sheetDimension[0] + 1];
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
                            columnWidths[columnWidthIndex] = RoundNumber(!double.IsNaN(columnWidthsTotal) ? columnWidths[columnWidthIndex] / columnWidthsTotal * 100 : columnWidths[columnWidthIndex] * 7, config.RoundingDigits);
                        }

                        int rowIndex = sheetDimension[1];
                        double[] rowHeightsAccumulation = new double[sheetDimension[3] - sheetDimension[1] + 1];
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
                                    rowHeightsAccumulation[additionalRowIndex - sheetDimension[1]] = (additionalRowIndex > sheetDimension[1] ? rowHeightsAccumulation[additionalRowIndex - sheetDimension[1] - 1] : 0) + (!double.IsNaN(rowHeightDefault) ? rowHeightDefault : 0);
                                    writer.Write($"\n{new string(' ', 12)}<tr>");
                                    for (int additionalColumnIndex = 0; additionalColumnIndex < columnWidths.Length; additionalColumnIndex++)
                                    {
                                        writer.Write($"\n{new string(' ', 16)}<td style=\"height: {rowHeightDefault}px; width: {(!double.IsNaN(columnWidths[additionalColumnIndex]) ? $"{columnWidths[additionalColumnIndex]}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")}" : "auto")};\"></td>");
                                    }
                                    writer.Write($"\n{new string(' ', 12)}</tr>");
                                }
                                rowIndex = (int)row.RowIndex.Value;
                            }
                            double cellHeightActual = config.ConvertSizes ? RoundNumber((row.CustomHeight == null || (row.CustomHeight.HasValue && row.CustomHeight.Value)) && row.Height != null && row.Height.HasValue ? row.Height.Value / 0.75 : rowHeightDefault, config.RoundingDigits) : double.NaN;
                            rowHeightsAccumulation[rowIndex - sheetDimension[1]] = (rowIndex > sheetDimension[1] ? rowHeightsAccumulation[rowIndex - sheetDimension[1] - 1] : 0) + (!double.IsNaN(cellHeightActual) ? cellHeightActual : 0);

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
                                bool isCellValid = true;
                                foreach (int[] mergeCellInfo in mergeCells)
                                {
                                    if (mergeCellInfo.Length > 5 && (mergeCellInfo[0] != columnIndex || mergeCellInfo[1] != rowIndex) && columnIndex >= mergeCellInfo[0] && columnIndex <= mergeCellInfo[2] && rowIndex >= mergeCellInfo[1] && rowIndex <= mergeCellInfo[3])
                                    {
                                        isCellValid = false;
                                        break;
                                    }
                                    else if (mergeCellInfo.Length > 5 && mergeCellInfo[0] == columnIndex && mergeCellInfo[1] == rowIndex)
                                    {
                                        columnSpanned = mergeCellInfo[4];
                                        rowSpanned = mergeCellInfo[5];
                                        break;
                                    }
                                }
                                if (!isCellValid)
                                {
                                    continue;
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
                                    foreach (ConditionalFormatting conditionalFormatting in conditionalFormattings)
                                    {
                                        if (conditionalFormatting.SequenceOfReferences != null && conditionalFormatting.SequenceOfReferences.HasValue)
                                        {
                                            bool isFormattingApplicable = false;
                                            foreach (string references in conditionalFormatting.SequenceOfReferences.Items)
                                            {
                                                int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
                                                int cellRowIndex = GetRowIndex(cell.CellReference.Value);
                                                GetReferenceRange(references, out int referenceFromColumn, out int referenceFromRow, out int referenceToColumn, out int referenceToRow);
                                                if (cellColumnIndex >= referenceFromColumn && cellColumnIndex <= referenceToColumn && cellRowIndex >= referenceFromRow && cellRowIndex <= referenceToRow)
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

                                            bool isConditionMet = false;
                                            if (formattingRule.Type.Value == ConditionalFormatValues.CellIs && formattingRule.Operator != null && formattingRule.Operator.HasValue)
                                            {
                                                switch (formattingRule.Operator.Value)
                                                {
                                                    case ConditionalFormattingOperatorValues.Equal:
                                                        isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaEqual && cellValueRaw == formulaEqual.Text.Trim('"');
                                                        break;
                                                    case ConditionalFormattingOperatorValues.NotEqual:
                                                        isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaNotEqual && cellValueRaw != formulaNotEqual.Text.Trim('"');
                                                        break;
                                                    case ConditionalFormattingOperatorValues.BeginsWith:
                                                        isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaBeginsWith && cellValueRaw.StartsWith(formulaBeginsWith.Text.Trim('"'));
                                                        break;
                                                    case ConditionalFormattingOperatorValues.EndsWith:
                                                        isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaEndsWith && cellValueRaw.EndsWith(formulaEndsWith.Text.Trim('"'));
                                                        break;
                                                    case ConditionalFormattingOperatorValues.ContainsText:
                                                        isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaContainsText && cellValueRaw.Contains(formulaContainsText.Text.Trim('"'));
                                                        break;
                                                    case ConditionalFormattingOperatorValues.NotContains:
                                                        isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaNotContains && !cellValueRaw.Contains(formulaNotContains.Text.Trim('"'));
                                                        break;
                                                    case ConditionalFormattingOperatorValues.GreaterThan:
                                                        isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] > x[1]);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.GreaterThanOrEqual:
                                                        isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] >= x[1]);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.LessThan:
                                                        isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] < x[1]);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.LessThanOrEqual:
                                                        isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] <= x[1]);
                                                        break;
                                                    case ConditionalFormattingOperatorValues.Between:
                                                        isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 2, x => x[0] >= Math.Min(x[1], x[2]) && x[0] <= Math.Max(x[1], x[2]));
                                                        break;
                                                    case ConditionalFormattingOperatorValues.NotBetween:
                                                        isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 2, x => x[0] < Math.Min(x[1], x[2]) || x[0] > Math.Max(x[1], x[2]));
                                                        break;
                                                }
                                            }
                                            else if (formattingRule.Type.Value == ConditionalFormatValues.BeginsWith && formattingRule.Text != null && formattingRule.Text.HasValue)
                                            {
                                                isConditionMet = cellValueRaw.StartsWith(formattingRule.Text.Value);
                                            }
                                            else if (formattingRule.Type.Value == ConditionalFormatValues.EndsWith && formattingRule.Text != null && formattingRule.Text.HasValue)
                                            {
                                                isConditionMet = cellValueRaw.EndsWith(formattingRule.Text.Value);
                                            }
                                            else if (formattingRule.Type.Value == ConditionalFormatValues.ContainsText && formattingRule.Text != null && formattingRule.Text.HasValue)
                                            {
                                                isConditionMet = cellValueRaw.Contains(formattingRule.Text.Value);
                                            }
                                            else if (formattingRule.Type.Value == ConditionalFormatValues.NotContainsText && formattingRule.Text != null && formattingRule.Text.HasValue)
                                            {
                                                isConditionMet = !cellValueRaw.Contains(formattingRule.Text.Value);
                                            }
                                            else if (formattingRule.Type.Value == ConditionalFormatValues.ContainsBlanks)
                                            {
                                                isConditionMet = string.IsNullOrWhiteSpace(cellValueRaw);
                                            }
                                            else if (formattingRule.Type.Value == ConditionalFormatValues.NotContainsBlanks)
                                            {
                                                isConditionMet = !string.IsNullOrWhiteSpace(cellValueRaw);
                                            }

                                            differentialStyleIndex = isConditionMet ? (int)formattingRule.FormatId.Value : differentialStyleIndex;
                                        }
                                    }
                                    if (differentialStyleIndex >= 0 && differentialStyleIndex < htmlStyleDifferentialFormats.Length && htmlStyleDifferentialFormats[differentialStyleIndex] != null)
                                    {
                                        htmlStyleCell = JoinHtmlAttributes(htmlStyleCell, htmlStyleDifferentialFormats[differentialStyleIndex].Item1);
                                        cellValueContainer = cellValueContainer.Replace("{0}", htmlStyleDifferentialFormats[differentialStyleIndex].Item2);
                                    }
                                }

                                writer.Write($"\n{new string(' ', 16)}<td{(columnSpanned > 1 ? $" colspan=\"{columnSpanned}\"" : string.Empty)}{(rowSpanned > 1 ? $" rowspan=\"{rowSpanned}\"" : string.Empty)} style=\"width: {(!double.IsNaN(cellWidthActual) && columnSpanned == 1 ? $"{cellWidthActual}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")}" : "auto")}; height: {(!double.IsNaN(cellHeightActual) && rowSpanned == 1 ? $"{cellHeightActual}px" : "auto")};{GetHtmlAttributesString(htmlStyleCell, true)}\">{cellValueContainer.Replace("{0}", cellValue)}</td>");
                            }

                            writer.Write($"\n{new string(' ', 12)}</tr>");

                            progressCallback?.Invoke(document, new ConverterProgressCallbackEventArgs(sheetIndex, sheetsCount, rowIndex, rowHeightsAccumulation.Length));
                        }

                        writer.Write($"\n{new string(' ', 8)}</table>");

                        if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                        {
                            //TODO: simplify, rounding, & better memory usage with row heights
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absoluteAnchor in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor>())
                            {
                                string left = absoluteAnchor.Position != null && absoluteAnchor.Position.X != null && absoluteAnchor.Position.X.HasValue ? $"{(double)absoluteAnchor.Position.X.Value / 914400.0 * 96}px" : "auto";
                                string top = absoluteAnchor.Position != null && absoluteAnchor.Position.Y != null && absoluteAnchor.Position.Y.HasValue ? $"{(double)absoluteAnchor.Position.Y.Value / 914400.0 * 96}px" : "auto";
                                string width = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cx != null && absoluteAnchor.Extent.Cx.HasValue ? $"{(double)absoluteAnchor.Extent.Cx.Value / 914400.0 * 96}px" : "auto";
                                string height = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cy != null && absoluteAnchor.Extent.Cy.HasValue ? $"{(double)absoluteAnchor.Extent.Cy.Value / 914400.0 * 96}px" : "auto";
                                DrawingsToHtml(worksheetPart, absoluteAnchor, writer, left, top, width, height, config);
                            }
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>())
                            {
                                double left = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) ? columnWidths.Take(Math.Min(columnWidths.Length, columnId - sheetDimension[0] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double leftOffset = oneCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(oneCellAnchor.FromMarker.ColumnOffset.Text, out double columnOffset) ? columnOffset / 914400.0 * 96 : 0;
                                double top = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) ? rowHeightsAccumulation[rowId - sheetDimension[1]] : double.NaN;
                                double topOffset = oneCellAnchor.FromMarker.RowOffset != null && double.TryParse(oneCellAnchor.FromMarker.RowOffset.Text, out double rowOffset) ? rowOffset / 914400.0 * 96 : 0;
                                string width = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cx != null && oneCellAnchor.Extent.Cx.HasValue ? $"{(double)oneCellAnchor.Extent.Cx.Value / 914400.0 * 96}px" : "auto";
                                string height = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cy != null && oneCellAnchor.Extent.Cy.HasValue ? $"{(double)oneCellAnchor.Extent.Cy.Value / 914400.0 * 96}px" : "auto";
                                DrawingsToHtml(worksheetPart, oneCellAnchor, writer, $"{(!double.IsNaN(left) ? $"calc({left}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")} + {leftOffset}px)" : $"{leftOffset}px")}", $"{(!double.IsNaN(top) ? top + topOffset : topOffset)}px", width, height, config);
                            }
                            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                            {
                                double fromColumn = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) ? columnWidths.Take(Math.Min(columnWidths.Length, fromColumnId - sheetDimension[0] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double fromColumnOffset = twoCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.FromMarker.ColumnOffset.Text, out double fromMarkerColumnOffset) ? fromMarkerColumnOffset / 914400.0 * 96 : 0;
                                double fromRow = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) ? rowHeightsAccumulation[fromRowId - sheetDimension[1]] : double.NaN;
                                double fromRowOffset = twoCellAnchor.FromMarker.RowOffset != null && double.TryParse(twoCellAnchor.FromMarker.RowOffset.Text, out double fromMarkerRowOffset) ? fromMarkerRowOffset / 914400.0 * 96 : 0;
                                double toColumn = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) ? columnWidths.Take(Math.Min(columnWidths.Length, toColumnId - sheetDimension[0] + 1)).SkipWhile(x => double.IsNaN(x)).Sum() : double.NaN;
                                double toColumnOffset = twoCellAnchor.ToMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.ToMarker.ColumnOffset.Text, out double toMarkerColumnOffset) ? toMarkerColumnOffset / 914400.0 * 96 : 0;
                                double toRow = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) ? rowHeightsAccumulation[toRowId - sheetDimension[1]] : double.NaN;
                                double toRowOffset = twoCellAnchor.ToMarker.RowOffset != null && double.TryParse(twoCellAnchor.ToMarker.RowOffset.Text, out double toMarkerRowOffset) ? toMarkerRowOffset / 914400.0 * 96 : 0;
                                string leftCalculation = $"{(!double.IsNaN(fromColumn) ? $"calc({fromColumn}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")} + {fromColumnOffset}px)" : $"{fromColumnOffset}px")}";
                                string topCalculation = $"{(!double.IsNaN(fromRow) ? fromRow + fromRowOffset : fromRowOffset)}px";
                                DrawingsToHtml(worksheetPart, twoCellAnchor, writer, leftCalculation, topCalculation, $"calc({(!double.IsNaN(toColumn) ? $"calc({toColumn}{(!double.IsNaN(columnWidthsTotal) ? "%" : "px")} + {toColumnOffset}px)" : $"{toColumnOffset}px")} - {leftCalculation})", $"calc({$"{(!double.IsNaN(toRow) ? toRow + toRowOffset : toRowOffset)}px"} - {topCalculation})", config);
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

        #region Private Fields

        private static readonly Regex regexNumbers = new Regex(@"\d+", RegexOptions.Compiled);
        private static readonly Regex regexLetters = new Regex("[A-Za-z]+", RegexOptions.Compiled);

        #endregion

        #region Private Methods

        private static void UpdateArray<T>(T[] array, params T[] values)
        {
            for (int i = 0; i < values.Length && i < array.Length; i++)
            {
                array[i] = values[i];
            }
        }

        private static double RoundNumber(double number, int digits)
        {
            return digits < 0 ? number : Math.Round(number, digits);
        }

        private static int GetColumnIndex(string cell)
        {
            int index = -1;
            Match match = regexLetters.Match(cell);
            if (match.Success)
            {
                int mulitplier = 1;
                foreach (char c in match.Value.ToUpper().ToCharArray().Reverse())
                {
                    index += mulitplier * (c - 64);
                    mulitplier *= 26;
                }
            }
            return Math.Max(0, index + 1);
        }

        private static int GetRowIndex(string cell)
        {
            Match match = regexNumbers.Match(cell);
            return match.Success && int.TryParse(match.Value, out int index) ? index : 0;
        }

        private static void GetReferenceRange(string range, out int fromColumn, out int fromRow, out int toColumn, out int toRow)
        {
            string[] rangeSplitted = range.Split(':');
            int firstColumn = GetColumnIndex(rangeSplitted[0]);
            int firstRow = GetRowIndex(rangeSplitted[0]);
            int secondColumn = rangeSplitted.Length > 1 ? GetColumnIndex(rangeSplitted[1]) : firstColumn;
            int secondRow = rangeSplitted.Length > 1 ? GetRowIndex(rangeSplitted[1]) : firstRow;
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
            if (attributes == null)
            {
                return string.Empty;
            }

            string htmlAttributes = string.Empty;
            foreach (KeyValuePair<string, string> pair in attributes)
            {
                if (!string.IsNullOrEmpty(pair.Key) && !string.IsNullOrEmpty(pair.Value))
                {
                    htmlAttributes += $"{pair.Key}: {pair.Value}; ";
                }
            }
            return isAdditional ? $" {htmlAttributes.TrimEnd()}" : htmlAttributes.TrimEnd();
        }

        private static Dictionary<string, string> JoinHtmlAttributes(Dictionary<string, string> original, Dictionary<string, string> joining)
        {
            if (joining == null)
            {
                return original;
            }

            foreach (KeyValuePair<string, string> pair in joining)
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

        private static bool GetNumberFormulaCondition(string value, IEnumerable<Formula> formulas, int formulasCount, Func<double[], bool> actionEvaluation)
        {
            if (!double.TryParse(value, out double valueDouble))
            {
                return false;
            }

            double[] parameters = new double[formulasCount + 1];
            parameters[0] = valueDouble;

            int index = 0;
            foreach (Formula formula in formulas)
            {
                index++;
                if (index > formulasCount || !double.TryParse(formula.Text, out double formulaDouble))
                {
                    break;
                }
                parameters[index] = formulaDouble;
            }
            return index >= formulasCount && actionEvaluation.Invoke(parameters);
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
                            Func<int, int, int, int, Func<int, int, bool>, bool> actionFindNewValue = (indexNumerator, indexDenominator, incrementNumerator, incrementDenominator, actionEvaluation) =>
                            {
                                fractionParts[indexNumerator] += incrementNumerator;
                                fractionParts[indexDenominator] += incrementDenominator;
                                if (actionEvaluation.Invoke(fractionParts[indexNumerator], fractionParts[indexDenominator]))
                                {
                                    int weight = 1;
                                    do
                                    {
                                        weight *= 2;
                                        fractionParts[indexNumerator] += incrementNumerator * weight;
                                        fractionParts[indexDenominator] += incrementDenominator * weight;
                                    }
                                    while (actionEvaluation.Invoke(fractionParts[indexNumerator], fractionParts[indexDenominator]));
                                    do
                                    {
                                        weight /= 2;
                                        int decrementNumerator = incrementNumerator * weight;
                                        int decrementDenominator = incrementDenominator * weight;
                                        if (!actionEvaluation.Invoke(fractionParts[indexNumerator] - decrementNumerator, fractionParts[indexDenominator] - decrementDenominator))
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
                                    actionFindNewValue.Invoke(2, 3, fractionParts[0], fractionParts[1], (numerator, denominator) => (fractionParts[1] + denominator) * (valueAbsolute + maxError) < (fractionParts[0] + numerator));
                                }
                                else if (middleNumerator < (valueAbsolute - maxError) * middleDenominator)
                                {
                                    actionFindNewValue.Invoke(0, 1, fractionParts[2], fractionParts[3], (numerator, denominator) => (numerator + fractionParts[2]) < (valueAbsolute - maxError) * (denominator + fractionParts[3]));
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

        private static Dictionary<string, string> CellFormatToHtml(WorkbookPart workbook, Fill fill, Font font, Border border, Alignment alignment, out string cellValueContainer, ConverterConfig config)
        {
            Dictionary<string, string> htmlStyle = new Dictionary<string, string>();
            cellValueContainer = "{0}";
            if (fill != null && fill.PatternFill != null && (fill.PatternFill.PatternType == null || (fill.PatternFill.PatternType.HasValue && fill.PatternFill.PatternType.Value != PatternValues.None)))
            {
                string background = string.Empty;
                if (fill.PatternFill.ForegroundColor != null)
                {
                    background = ColorTypeToHtml(workbook, fill.PatternFill.ForegroundColor, config);
                }
                if (string.IsNullOrEmpty(background) && fill.PatternFill.BackgroundColor != null)
                {
                    background = ColorTypeToHtml(workbook, fill.PatternFill.BackgroundColor, config);
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
                htmlStyle = JoinHtmlAttributes(htmlStyle, FontToHtml(workbook, font.Color, font.FontSize, font.Bold, font.Italic, font.Strike, font.Underline, out cellValueContainer, config));
            }
            if (border != null)
            {
                if (border.LeftBorder != null)
                {
                    string borderLeft = BorderPropertiesToHtml(workbook, border.LeftBorder, config);
                    if (!string.IsNullOrEmpty(borderLeft))
                    {
                        htmlStyle.Add("border-left", borderLeft);
                    }
                }
                if (border.RightBorder != null)
                {
                    string borderRight = BorderPropertiesToHtml(workbook, border.RightBorder, config);
                    if (!string.IsNullOrEmpty(borderRight))
                    {
                        htmlStyle.Add("border-right", borderRight);
                    }
                }
                if (border.TopBorder != null)
                {
                    string borderTop = BorderPropertiesToHtml(workbook, border.TopBorder, config);
                    if (!string.IsNullOrEmpty(borderTop))
                    {
                        htmlStyle.Add("border-top", borderTop);
                    }
                }
                if (border.BottomBorder != null)
                {
                    string borderBottom = BorderPropertiesToHtml(workbook, border.BottomBorder, config);
                    if (!string.IsNullOrEmpty(borderBottom))
                    {
                        htmlStyle.Add("border-bottom", borderBottom);
                    }
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
                    cellValueContainer = cellValueContainer.Replace("{0}", $"<div style=\"width: fit-content; transform: rotate(-{RoundNumber(alignment.TextRotation.Value, config.RoundingDigits)}deg);\">{{0}}</div>");
                }
            }
            return htmlStyle;
        }

        private static Dictionary<string, string> FontToHtml(WorkbookPart workbook, ColorType color, FontSize fontSize, Bold bold, Italic italic, Strike strike, Underline underline, out string cellValueContainer, ConverterConfig config)
        {
            Dictionary<string, string> htmlStyle = new Dictionary<string, string>();
            cellValueContainer = "{0}";
            if (color != null)
            {
                string htmlColor = ColorTypeToHtml(workbook, color, config);
                if (!string.IsNullOrEmpty(htmlColor))
                {
                    htmlStyle.Add("color", htmlColor);
                }
            }
            if (fontSize != null && fontSize.Val != null && fontSize.Val.HasValue)
            {
                htmlStyle.Add("font-size", $"{RoundNumber(fontSize.Val.Value / 72 * 96, config.RoundingDigits)}px");
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
                    cellValueContainer = cellValueContainer.Replace("{0}", $"<div style=\"width: fit-content; text-decoration: underline double;\">{{0}}</div>");
                }
                else if (underline.Val.Value != UnderlineValues.None)
                {
                    htmlStyleTextDecoraion += " underline";
                }
            }
            if (!string.IsNullOrEmpty(htmlStyleTextDecoraion))
            {
                htmlStyle.Add("text-decoration", htmlStyleTextDecoraion.TrimStart());
            }
            return htmlStyle;
        }

        private static string BorderPropertiesToHtml(WorkbookPart workbook, BorderPropertiesType border, ConverterConfig config)
        {
            if (border == null)
            {
                return string.Empty;
            }

            string htmlBorder = string.Empty;
            if (border.Style != null && border.Style.HasValue)
            {
                switch (border.Style.Value)
                {
                    case BorderStyleValues.Hair:
                        htmlBorder += " thin solid";
                        break;
                    case BorderStyleValues.Thin:
                        htmlBorder += " thin solid";
                        break;
                    case BorderStyleValues.Thick:
                        htmlBorder += " thick solid";
                        break;
                    case BorderStyleValues.Medium:
                        htmlBorder += " medium solid";
                        break;
                    case BorderStyleValues.MediumDashDot:
                        htmlBorder += " medium dashed";
                        break;
                    case BorderStyleValues.MediumDashDotDot:
                        htmlBorder += " medium dotted";
                        break;
                    case BorderStyleValues.MediumDashed:
                        htmlBorder += " medium dashed";
                        break;
                    case BorderStyleValues.Dashed:
                        htmlBorder += " dashed";
                        break;
                    case BorderStyleValues.DashDot:
                        htmlBorder += " dashed";
                        break;
                    case BorderStyleValues.DashDotDot:
                        htmlBorder += " dotted";
                        break;
                    case BorderStyleValues.Dotted:
                        htmlBorder += " dotted";
                        break;
                    case BorderStyleValues.SlantDashDot:
                        htmlBorder += " dashed";
                        break;
                    case BorderStyleValues.Double:
                        htmlBorder += " double";
                        break;
                }
            }
            if (border.Color != null)
            {
                string value = ColorTypeToHtml(workbook, border.Color, config);
                if (!string.IsNullOrEmpty(value))
                {
                    htmlBorder += $" {value}";
                }
            }
            return htmlBorder.TrimStart();
        }

        private static string ColorTypeToHtml(WorkbookPart workbook, ColorType color, ConverterConfig config)
        {
            if (color == null)
            {
                return string.Empty;
            }

            double[] result = new double[4] { 0, 0, 0, 1 };
            if (color.Auto != null && color.Auto.HasValue && color.Auto.Value)
            {
                return "initial";
            }
            else if (color.Rgb != null && color.Rgb.HasValue)
            {
                HexToRgba(color.Rgb.Value, out result[0], out result[1], out result[2], out result[3]);
            }
            else if (color.Indexed != null && color.Indexed.HasValue)
            {
                switch (color.Indexed.Value)
                {
                    case 0:
                        UpdateArray(result, 0, 0, 0);
                        break;
                    case 1:
                        UpdateArray(result, 255, 255, 255);
                        break;
                    case 2:
                        UpdateArray(result, 255, 0, 0);
                        break;
                    case 3:
                        UpdateArray(result, 0, 255, 0);
                        break;
                    case 4:
                        UpdateArray(result, 0, 0, 255);
                        break;
                    case 5:
                        UpdateArray(result, 255, 255, 0);
                        break;
                    case 6:
                        UpdateArray(result, 255, 0, 255);
                        break;
                    case 7:
                        UpdateArray(result, 0, 255, 255);
                        break;
                    case 8:
                        UpdateArray(result, 0, 0, 0);
                        break;
                    case 9:
                        UpdateArray(result, 255, 255, 255);
                        break;
                    case 10:
                        UpdateArray(result, 255, 0, 0);
                        break;
                    case 11:
                        UpdateArray(result, 0, 255, 0);
                        break;
                    case 12:
                        UpdateArray(result, 0, 0, 255);
                        break;
                    case 13:
                        UpdateArray(result, 255, 255, 0);
                        break;
                    case 14:
                        UpdateArray(result, 255, 0, 255);
                        break;
                    case 15:
                        UpdateArray(result, 0, 255, 255);
                        break;
                    case 16:
                        UpdateArray(result, 128, 0, 0);
                        break;
                    case 17:
                        UpdateArray(result, 0, 128, 0);
                        break;
                    case 18:
                        UpdateArray(result, 0, 0, 128);
                        break;
                    case 19:
                        UpdateArray(result, 128, 128, 0);
                        break;
                    case 20:
                        UpdateArray(result, 128, 0, 128);
                        break;
                    case 21:
                        UpdateArray(result, 0, 128, 128);
                        break;
                    case 22:
                        UpdateArray(result, 192, 192, 192);
                        break;
                    case 23:
                        UpdateArray(result, 128, 128, 128);
                        break;
                    case 24:
                        UpdateArray(result, 153, 153, 255);
                        break;
                    case 25:
                        UpdateArray(result, 153, 51, 102);
                        break;
                    case 26:
                        UpdateArray(result, 255, 255, 204);
                        break;
                    case 27:
                        UpdateArray(result, 204, 255, 255);
                        break;
                    case 28:
                        UpdateArray(result, 102, 0, 102);
                        break;
                    case 29:
                        UpdateArray(result, 255, 128, 128);
                        break;
                    case 30:
                        UpdateArray(result, 0, 102, 204);
                        break;
                    case 31:
                        UpdateArray(result, 204, 204, 255);
                        break;
                    case 32:
                        UpdateArray(result, 0, 0, 128);
                        break;
                    case 33:
                        UpdateArray(result, 255, 0, 255);
                        break;
                    case 34:
                        UpdateArray(result, 255, 255, 0);
                        break;
                    case 35:
                        UpdateArray(result, 0, 255, 255);
                        break;
                    case 36:
                        UpdateArray(result, 128, 0, 128);
                        break;
                    case 37:
                        UpdateArray(result, 128, 0, 0);
                        break;
                    case 38:
                        UpdateArray(result, 0, 128, 128);
                        break;
                    case 39:
                        UpdateArray(result, 0, 0, 255);
                        break;
                    case 40:
                        UpdateArray(result, 0, 204, 255);
                        break;
                    case 41:
                        UpdateArray(result, 204, 255, 255);
                        break;
                    case 42:
                        UpdateArray(result, 204, 255, 204);
                        break;
                    case 43:
                        UpdateArray(result, 255, 255, 153);
                        break;
                    case 44:
                        UpdateArray(result, 153, 204, 255);
                        break;
                    case 45:
                        UpdateArray(result, 255, 153, 204);
                        break;
                    case 46:
                        UpdateArray(result, 204, 153, 255);
                        break;
                    case 47:
                        UpdateArray(result, 255, 204, 153);
                        break;
                    case 48:
                        UpdateArray(result, 51, 102, 255);
                        break;
                    case 49:
                        UpdateArray(result, 51, 204, 204);
                        break;
                    case 50:
                        UpdateArray(result, 153, 204, 0);
                        break;
                    case 51:
                        UpdateArray(result, 255, 204, 0);
                        break;
                    case 52:
                        UpdateArray(result, 255, 153, 0);
                        break;
                    case 53:
                        UpdateArray(result, 255, 102, 0);
                        break;
                    case 54:
                        UpdateArray(result, 102, 102, 153);
                        break;
                    case 55:
                        UpdateArray(result, 150, 150, 150);
                        break;
                    case 56:
                        UpdateArray(result, 0, 51, 102);
                        break;
                    case 57:
                        UpdateArray(result, 51, 153, 102);
                        break;
                    case 58:
                        UpdateArray(result, 0, 51, 0);
                        break;
                    case 59:
                        UpdateArray(result, 51, 51, 0);
                        break;
                    case 60:
                        UpdateArray(result, 153, 51, 0);
                        break;
                    case 61:
                        UpdateArray(result, 153, 51, 102);
                        break;
                    case 62:
                        UpdateArray(result, 51, 51, 153);
                        break;
                    case 63:
                        UpdateArray(result, 51, 51, 51);
                        break;
                    case 64:
                        UpdateArray(result, 128, 128, 128);
                        break;
                    case 65:
                        UpdateArray(result, 255, 255, 255);
                        break;
                    default:
                        return string.Empty;
                }
            }
            else if (color.Theme != null && color.Theme.HasValue && workbook.ThemePart != null && workbook.ThemePart.Theme != null && workbook.ThemePart.Theme.ThemeElements != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme != null)
            {
                DocumentFormat.OpenXml.Drawing.Color2Type themeColor = null;
                switch (color.Theme.Value)
                {
                    case 0:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Light1Color;
                        break;
                    case 1:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Dark1Color;
                        break;
                    case 2:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Light2Color;
                        break;
                    case 3:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Dark2Color;
                        break;
                    case 4:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent1Color;
                        break;
                    case 5:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent2Color;
                        break;
                    case 6:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent3Color;
                        break;
                    case 7:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent4Color;
                        break;
                    case 8:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent5Color;
                        break;
                    case 9:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent6Color;
                        break;
                    case 10:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.Hyperlink;
                        break;
                    case 11:
                        themeColor = workbook.ThemePart.Theme.ThemeElements.ColorScheme.FollowedHyperlinkColor;
                        break;
                }

                if (themeColor != null && themeColor.RgbColorModelHex != null && themeColor.RgbColorModelHex.Val != null && themeColor.RgbColorModelHex.Val.HasValue)
                {
                    HexToRgba(themeColor.RgbColorModelHex.Val.Value, out result[0], out result[1], out result[2], out result[3]);
                }
                else if (themeColor != null && themeColor.RgbColorModelPercentage != null)
                {
                    result[0] = themeColor.RgbColorModelPercentage.RedPortion != null && themeColor.RgbColorModelPercentage.RedPortion.HasValue ? (int)(themeColor.RgbColorModelPercentage.RedPortion.Value / 100000.0 * 255) : 0;
                    result[1] = themeColor.RgbColorModelPercentage.GreenPortion != null && themeColor.RgbColorModelPercentage.GreenPortion.HasValue ? (int)(themeColor.RgbColorModelPercentage.GreenPortion.Value / 100000.0 * 255) : 0;
                    result[2] = themeColor.RgbColorModelPercentage.BluePortion != null && themeColor.RgbColorModelPercentage.BluePortion.HasValue ? (int)(themeColor.RgbColorModelPercentage.BluePortion.Value / 100000.0 * 255) : 0;
                }
                else if (themeColor != null && themeColor.HslColor != null)
                {
                    double hue = themeColor.HslColor.HueValue != null && themeColor.HslColor.HueValue.HasValue ? themeColor.HslColor.HueValue.Value / 60000.0 : 0;
                    double luminosity = themeColor.HslColor.LumValue != null && themeColor.HslColor.LumValue.HasValue ? themeColor.HslColor.LumValue.Value / 100000.0 : 0;
                    double saturation = themeColor.HslColor.SatValue != null && themeColor.HslColor.SatValue.HasValue ? themeColor.HslColor.SatValue.Value / 100000.0 : 0;
                    HlsToRgb(hue, luminosity, saturation, out result[0], out result[1], out result[2]);
                }
                else if (themeColor != null && themeColor.SystemColor != null)
                {
                    if (themeColor.SystemColor.Val != null && themeColor.SystemColor.Val.HasValue)
                    {
                        switch (themeColor.SystemColor.Val.Value)
                        {
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveBorder:
                                UpdateArray(result, 180, 180, 180);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveCaption:
                                UpdateArray(result, 153, 180, 209);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ApplicationWorkspace:
                                UpdateArray(result, 171, 171, 171);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.Background:
                                UpdateArray(result, 255, 255, 255);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonFace:
                                UpdateArray(result, 240, 240, 240);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonHighlight:
                                UpdateArray(result, 0, 120, 215);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonShadow:
                                UpdateArray(result, 160, 160, 160);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonText:
                                UpdateArray(result, 0, 0, 0);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.CaptionText:
                                UpdateArray(result, 0, 0, 0);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientActiveCaption:
                                UpdateArray(result, 185, 209, 234);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientInactiveCaption:
                                UpdateArray(result, 215, 228, 242);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.GrayText:
                                UpdateArray(result, 109, 109, 109);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.Highlight:
                                UpdateArray(result, 0, 120, 215);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.HighlightText:
                                UpdateArray(result, 255, 255, 255);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.HotLight:
                                UpdateArray(result, 255, 165, 0);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveBorder:
                                UpdateArray(result, 244, 247, 252);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaption:
                                UpdateArray(result, 191, 205, 219);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaptionText:
                                UpdateArray(result, 0, 0, 0);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoBack:
                                UpdateArray(result, 255, 255, 225);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoText:
                                UpdateArray(result, 0, 0, 0);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.Menu:
                                UpdateArray(result, 240, 240, 240);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuBar:
                                UpdateArray(result, 240, 240, 240);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuHighlight:
                                UpdateArray(result, 0, 120, 215);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuText:
                                UpdateArray(result, 0, 0, 0);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ScrollBar:
                                UpdateArray(result, 200, 200, 200);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDDarkShadow:
                                UpdateArray(result, 160, 160, 160);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDLight:
                                UpdateArray(result, 227, 227, 227);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.Window:
                                UpdateArray(result, 255, 255, 255);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowFrame:
                                UpdateArray(result, 100, 100, 100);
                                break;
                            case DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText:
                                UpdateArray(result, 0, 0, 0);
                                break;
                            default:
                                return string.Empty;
                        };
                    }
                    else if (themeColor.SystemColor.LastColor != null && themeColor.SystemColor.LastColor.HasValue)
                    {
                        HexToRgba(themeColor.SystemColor.LastColor.Value, out result[0], out result[1], out result[2], out result[3]);
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
                else if (themeColor != null && themeColor.PresetColor != null && themeColor.PresetColor.Val != null && themeColor.PresetColor.Val.HasValue)
                {
                    switch (themeColor.PresetColor.Val.Value)
                    {
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.AliceBlue:
                            UpdateArray(result, 240, 248, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.AntiqueWhite:
                            UpdateArray(result, 250, 235, 215);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Aqua:
                            UpdateArray(result, 0, 255, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Aquamarine:
                            UpdateArray(result, 127, 255, 212);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Azure:
                            UpdateArray(result, 240, 255, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Beige:
                            UpdateArray(result, 245, 245, 220);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Bisque:
                            UpdateArray(result, 255, 228, 196);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Black:
                            UpdateArray(result, 0, 0, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.BlanchedAlmond:
                            UpdateArray(result, 255, 235, 205);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Blue:
                            UpdateArray(result, 0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.BlueViolet:
                            UpdateArray(result, 138, 43, 226);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Brown:
                            UpdateArray(result, 165, 42, 42);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.BurlyWood:
                            UpdateArray(result, 222, 184, 135);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.CadetBlue:
                            UpdateArray(result, 95, 158, 160);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Chartreuse:
                            UpdateArray(result, 127, 255, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Chocolate:
                            UpdateArray(result, 210, 105, 30);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Coral:
                            UpdateArray(result, 255, 127, 80);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.CornflowerBlue:
                            UpdateArray(result, 100, 149, 237);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Cornsilk:
                            UpdateArray(result, 255, 248, 220);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Crimson:
                            UpdateArray(result, 220, 20, 60);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Cyan:
                            UpdateArray(result, 0, 255, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue:
                            UpdateArray(result, 0, 0, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan:
                            UpdateArray(result, 0, 139, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod:
                            UpdateArray(result, 184, 134, 11);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray:
                            UpdateArray(result, 169, 169, 169);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen:
                            UpdateArray(result, 0, 100, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki:
                            UpdateArray(result, 189, 183, 107);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta:
                            UpdateArray(result, 139, 0, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen:
                            UpdateArray(result, 85, 107, 47);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange:
                            UpdateArray(result, 255, 140, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid:
                            UpdateArray(result, 153, 50, 204);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed:
                            UpdateArray(result, 139, 0, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon:
                            UpdateArray(result, 233, 150, 122);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen:
                            UpdateArray(result, 143, 188, 143);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue:
                            UpdateArray(result, 72, 61, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray:
                            UpdateArray(result, 47, 79, 79);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise:
                            UpdateArray(result, 0, 206, 209);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet:
                            UpdateArray(result, 148, 0, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepPink:
                            UpdateArray(result, 255, 20, 147);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepSkyBlue:
                            UpdateArray(result, 0, 191, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGray:
                            UpdateArray(result, 105, 105, 105);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DodgerBlue:
                            UpdateArray(result, 30, 144, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Firebrick:
                            UpdateArray(result, 178, 34, 34);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.FloralWhite:
                            UpdateArray(result, 255, 250, 240);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.ForestGreen:
                            UpdateArray(result, 34, 139, 34);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Fuchsia:
                            UpdateArray(result, 255, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Gainsboro:
                            UpdateArray(result, 220, 220, 220);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.GhostWhite:
                            UpdateArray(result, 248, 248, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Gold:
                            UpdateArray(result, 255, 215, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Goldenrod:
                            UpdateArray(result, 218, 165, 32);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Gray:
                            UpdateArray(result, 128, 128, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Green:
                            UpdateArray(result, 0, 128, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.GreenYellow:
                            UpdateArray(result, 173, 255, 47);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Honeydew:
                            UpdateArray(result, 240, 255, 240);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.HotPink:
                            UpdateArray(result, 255, 105, 180);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.IndianRed:
                            UpdateArray(result, 205, 92, 92);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Indigo:
                            UpdateArray(result, 75, 0, 130);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Ivory:
                            UpdateArray(result, 255, 255, 240);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Khaki:
                            UpdateArray(result, 240, 230, 140);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Lavender:
                            UpdateArray(result, 230, 230, 250);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LavenderBlush:
                            UpdateArray(result, 255, 240, 245);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LawnGreen:
                            UpdateArray(result, 124, 252, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LemonChiffon:
                            UpdateArray(result, 255, 250, 205);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue:
                            UpdateArray(result, 173, 216, 230);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral:
                            UpdateArray(result, 240, 128, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan:
                            UpdateArray(result, 224, 255, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow:
                            UpdateArray(result, 250, 250, 210);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray:
                            UpdateArray(result, 211, 211, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen:
                            UpdateArray(result, 144, 238, 144);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink:
                            UpdateArray(result, 255, 182, 193);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon:
                            UpdateArray(result, 255, 160, 122);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen:
                            UpdateArray(result, 32, 178, 170);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue:
                            UpdateArray(result, 135, 206, 250);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray:
                            UpdateArray(result, 119, 136, 153);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue:
                            UpdateArray(result, 176, 196, 222);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow:
                            UpdateArray(result, 255, 255, 224);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Lime:
                            UpdateArray(result, 0, 255, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LimeGreen:
                            UpdateArray(result, 50, 205, 50);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Linen:
                            UpdateArray(result, 250, 240, 230);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Magenta:
                            UpdateArray(result, 255, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Maroon:
                            UpdateArray(result, 128, 0, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MedAquamarine:
                            UpdateArray(result, 102, 205, 170);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue:
                            UpdateArray(result, 0, 0, 205);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid:
                            UpdateArray(result, 186, 85, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple:
                            UpdateArray(result, 147, 112, 219);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen:
                            UpdateArray(result, 60, 179, 113);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue:
                            UpdateArray(result, 123, 104, 238);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen:
                            UpdateArray(result, 0, 250, 154);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise:
                            UpdateArray(result, 72, 209, 204);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed:
                            UpdateArray(result, 199, 21, 133);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MidnightBlue:
                            UpdateArray(result, 25, 25, 112);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MintCream:
                            UpdateArray(result, 245, 255, 250);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MistyRose:
                            UpdateArray(result, 255, 228, 225);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Moccasin:
                            UpdateArray(result, 255, 228, 181);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.NavajoWhite:
                            UpdateArray(result, 255, 222, 173);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Navy:
                            UpdateArray(result, 0, 0, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.OldLace:
                            UpdateArray(result, 253, 245, 230);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Olive:
                            UpdateArray(result, 128, 128, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.OliveDrab:
                            UpdateArray(result, 107, 142, 35);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Orange:
                            UpdateArray(result, 255, 165, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.OrangeRed:
                            UpdateArray(result, 255, 69, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Orchid:
                            UpdateArray(result, 218, 112, 214);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGoldenrod:
                            UpdateArray(result, 238, 232, 170);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGreen:
                            UpdateArray(result, 152, 251, 152);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleTurquoise:
                            UpdateArray(result, 175, 238, 238);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleVioletRed:
                            UpdateArray(result, 219, 112, 147);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PapayaWhip:
                            UpdateArray(result, 255, 239, 213);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PeachPuff:
                            UpdateArray(result, 255, 218, 185);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Peru:
                            UpdateArray(result, 205, 133, 63);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Pink:
                            UpdateArray(result, 255, 192, 203);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Plum:
                            UpdateArray(result, 221, 160, 221);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.PowderBlue:
                            UpdateArray(result, 176, 224, 230);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Purple:
                            UpdateArray(result, 128, 0, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Red:
                            UpdateArray(result, 255, 0, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.RosyBrown:
                            UpdateArray(result, 188, 143, 143);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.RoyalBlue:
                            UpdateArray(result, 65, 105, 225);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SaddleBrown:
                            UpdateArray(result, 139, 69, 19);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Salmon:
                            UpdateArray(result, 250, 128, 114);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SandyBrown:
                            UpdateArray(result, 244, 164, 96);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaGreen:
                            UpdateArray(result, 46, 139, 87);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaShell:
                            UpdateArray(result, 255, 245, 238);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Sienna:
                            UpdateArray(result, 160, 82, 45);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Silver:
                            UpdateArray(result, 192, 192, 192);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SkyBlue:
                            UpdateArray(result, 135, 206, 235);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateBlue:
                            UpdateArray(result, 106, 90, 205);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGray:
                            UpdateArray(result, 112, 128, 144);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Snow:
                            UpdateArray(result, 255, 250, 250);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SpringGreen:
                            UpdateArray(result, 0, 255, 127);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SteelBlue:
                            UpdateArray(result, 70, 130, 180);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Tan:
                            UpdateArray(result, 210, 180, 140);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Teal:
                            UpdateArray(result, 0, 128, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Thistle:
                            UpdateArray(result, 216, 191, 216);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Tomato:
                            UpdateArray(result, 255, 99, 71);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Turquoise:
                            UpdateArray(result, 64, 224, 208);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Violet:
                            UpdateArray(result, 238, 130, 238);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Wheat:
                            UpdateArray(result, 245, 222, 179);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.White:
                            UpdateArray(result, 255, 255, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.WhiteSmoke:
                            UpdateArray(result, 245, 245, 245);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Yellow:
                            UpdateArray(result, 255, 255, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.YellowGreen:
                            UpdateArray(result, 154, 205, 50);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue2010:
                            UpdateArray(result, 0, 0, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan2010:
                            UpdateArray(result, 0, 139, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod2010:
                            UpdateArray(result, 184, 134, 11);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray2010:
                            UpdateArray(result, 169, 169, 169);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey2010:
                            UpdateArray(result, 169, 169, 169);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen2010:
                            UpdateArray(result, 0, 100, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki2010:
                            UpdateArray(result, 189, 183, 107);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta2010:
                            UpdateArray(result, 139, 0, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen2010:
                            UpdateArray(result, 85, 107, 47);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange2010:
                            UpdateArray(result, 255, 140, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid2010:
                            UpdateArray(result, 153, 50, 204);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed2010:
                            UpdateArray(result, 139, 0, 0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon2010:
                            UpdateArray(result, 233, 150, 122);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen2010:
                            UpdateArray(result, 143, 188, 143);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue2010:
                            UpdateArray(result, 72, 61, 139);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray2010:
                            UpdateArray(result, 47, 79, 79);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey2010:
                            UpdateArray(result, 47, 79, 79);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise2010:
                            UpdateArray(result, 0, 206, 209);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet2010:
                            UpdateArray(result, 148, 0, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue2010:
                            UpdateArray(result, 173, 216, 230);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral2010:
                            UpdateArray(result, 240, 128, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan2010:
                            UpdateArray(result, 224, 255, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow2010:
                            UpdateArray(result, 250, 250, 210);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray2010:
                            UpdateArray(result, 211, 211, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey2010:
                            UpdateArray(result, 211, 211, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen2010:
                            UpdateArray(result, 144, 238, 144);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink2010:
                            UpdateArray(result, 255, 182, 193);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon2010:
                            UpdateArray(result, 255, 160, 122);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen2010:
                            UpdateArray(result, 32, 178, 170);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue2010:
                            UpdateArray(result, 135, 206, 250);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray2010:
                            UpdateArray(result, 119, 136, 153);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey2010:
                            UpdateArray(result, 119, 136, 153);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue2010:
                            UpdateArray(result, 176, 196, 222);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow2010:
                            UpdateArray(result, 255, 255, 224);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumAquamarine2010:
                            UpdateArray(result, 102, 205, 170);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue2010:
                            UpdateArray(result, 0, 0, 205);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid2010:
                            UpdateArray(result, 186, 85, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple2010:
                            UpdateArray(result, 147, 112, 219);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen2010:
                            UpdateArray(result, 60, 179, 113);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue2010:
                            UpdateArray(result, 123, 104, 238);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen2010:
                            UpdateArray(result, 0, 250, 154);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise2010:
                            UpdateArray(result, 72, 209, 204);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed2010:
                            UpdateArray(result, 199, 21, 133);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey:
                            UpdateArray(result, 169, 169, 169);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGrey:
                            UpdateArray(result, 105, 105, 105);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey:
                            UpdateArray(result, 47, 79, 79);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.Grey:
                            UpdateArray(result, 128, 128, 128);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey:
                            UpdateArray(result, 211, 211, 211);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey:
                            UpdateArray(result, 119, 136, 153);
                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGrey:
                            UpdateArray(result, 112, 128, 144);
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

            if (color.Tint != null && color.Tint.HasValue && color.Tint.Value != 0)
            {
                RgbToHls(result[0], result[1], result[2], out double hue, out double luminosity, out double saturation);
                luminosity = color.Tint.Value < 0 ? luminosity * (1 + color.Tint.Value) : luminosity * (1 - color.Tint.Value) + color.Tint.Value;
                HlsToRgb(hue, luminosity, saturation, out result[0], out result[1], out result[2]);
            }

            return result.Length < 4 ? string.Empty : result[3] >= 1 ? $"rgb({RoundNumber(result[0], config.RoundingDigits)}, {RoundNumber(result[1], config.RoundingDigits)}, {RoundNumber(result[2], config.RoundingDigits)})" : $"rgba({RoundNumber(result[0], config.RoundingDigits)}, {RoundNumber(result[1], config.RoundingDigits)}, {RoundNumber(result[2], config.RoundingDigits)}, {RoundNumber(result[3], config.RoundingDigits)})";
        }

        private static void HexToRgba(string hex, out double red, out double green, out double blue, out double alpha)
        {
            string hexTrimmed = hex.TrimStart('#');
            red = hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 2 : 0, 2), 16) : 0;
            green = hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 4 : 2, 2), 16) : 0;
            blue = hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 6 : 4, 2), 16) : 0;
            alpha = hexTrimmed.Length > 5 ? (hexTrimmed.Length > 7 ? Convert.ToInt32(hexTrimmed.Substring(0, 2), 16) / 255.0 : 1) : 1;
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

        private static void DrawingsToHtml(WorksheetPart worksheet, OpenXmlCompositeElement anchor, StreamWriter writer, string left, string top, string width, string height, ConverterConfig config)
        {
            if (anchor == null)
            {
                return;
            }

            List<object[]> drawings = new List<object[]>();
            if (config.ConvertPictures)
            {
                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>())
                {
                    if (picture.BlipFill == null || picture.BlipFill.Blip == null)
                    {
                        continue;
                    }

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties nonVisualProperties = null;
                    if (picture.NonVisualPictureProperties != null)
                    {
                        nonVisualProperties = picture.NonVisualPictureProperties.NonVisualDrawingProperties;
                    }

                    if (picture.BlipFill.Blip.Embed != null && picture.BlipFill.Blip.Embed.HasValue && worksheet.DrawingsPart != null && worksheet.DrawingsPart.GetPartById(picture.BlipFill.Blip.Embed.Value) is ImagePart imagePart)
                    {
                        Stream imageStream = imagePart.GetStream();
                        if (!imageStream.CanRead)
                        {
                            continue;
                        }
                        else if (imageStream.CanSeek)
                        {
                            imageStream.Seek(0, SeekOrigin.Begin);
                        }
                        byte[] data = new byte[imageStream.Length];
                        imageStream.Read(data, 0, (int)imageStream.Length);

                        string base64 = Convert.ToBase64String(data, Base64FormattingOptions.None);
                        drawings.Add(new object[3] { $"<img src=\"data:{imagePart.ContentType};base64,{base64}\"{(nonVisualProperties != null && nonVisualProperties.Description != null && nonVisualProperties.Description.HasValue ? $" alt=\"{nonVisualProperties.Description.Value}\"" : string.Empty)}{{0}}/>", nonVisualProperties, picture.ShapeProperties });
                    }
                }
            }
            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>())
            {
                //TODO: shape styles
                string text = shape.TextBody != null ? shape.TextBody.InnerText : string.Empty;
                drawings.Add(new object[3] { $"<p{{0}}>{text}</p>", shape.NonVisualShapeProperties, shape.ShapeProperties });
            }

            foreach (object[] drawingInfo in drawings)
            {
                if (drawingInfo.Length < 3)
                {
                    continue;
                }

                string widthActual = width;
                string heightActual = height;
                string htmlStyleTransform = string.Empty;
                if (drawingInfo[2] is DocumentFormat.OpenXml.Drawing.ShapeProperties shapeProperties && shapeProperties.Transform2D != null)
                {
                    if (shapeProperties.Transform2D.Offset != null)
                    {
                        if (left == "auto" && shapeProperties.Transform2D.Offset.X != null && shapeProperties.Transform2D.Offset.X.HasValue)
                        {
                            htmlStyleTransform += $" translateX({RoundNumber(shapeProperties.Transform2D.Offset.X.Value / 914400.0 * 96, config.RoundingDigits)}px)";
                        }
                        if (top == "auto" && shapeProperties.Transform2D.Offset.Y != null && shapeProperties.Transform2D.Offset.Y.HasValue)
                        {
                            htmlStyleTransform += $" translateY({RoundNumber(shapeProperties.Transform2D.Offset.Y.Value / 914400.0 * 96, config.RoundingDigits)}px)";
                        }
                    }
                    if (shapeProperties.Transform2D.Extents != null)
                    {
                        if (widthActual == "auto" && shapeProperties.Transform2D.Extents.Cx != null && shapeProperties.Transform2D.Extents.Cx.HasValue)
                        {
                            widthActual = $"{RoundNumber(shapeProperties.Transform2D.Extents.Cx.Value / 914400.0 * 96, config.RoundingDigits)}px";
                        }
                        if (heightActual == "auto" && shapeProperties.Transform2D.Extents.Cy != null && shapeProperties.Transform2D.Extents.Cy.HasValue)
                        {
                            heightActual = $"{RoundNumber(shapeProperties.Transform2D.Extents.Cy.Value / 914400.0 * 96, config.RoundingDigits)}px";
                        }
                    }
                    if (shapeProperties.Transform2D.Rotation != null && shapeProperties.Transform2D.Rotation.HasValue)
                    {
                        htmlStyleTransform += $" rotate(-{RoundNumber(shapeProperties.Transform2D.Rotation.Value, config.RoundingDigits)}deg)";
                    }
                    if (shapeProperties.Transform2D.HorizontalFlip != null && shapeProperties.Transform2D.HorizontalFlip.HasValue && shapeProperties.Transform2D.HorizontalFlip.Value)
                    {
                        htmlStyleTransform += $" scaleX(-1)";
                    }
                    if (shapeProperties.Transform2D.VerticalFlip != null && shapeProperties.Transform2D.VerticalFlip.HasValue && shapeProperties.Transform2D.VerticalFlip.Value)
                    {
                        htmlStyleTransform += $" scaleY(-1)";
                    }
                }

                bool isHidden = drawingInfo[1] is DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties nonVisualProperties && nonVisualProperties.Hidden != null && nonVisualProperties.Hidden.HasValue && nonVisualProperties.Hidden.Value;
                writer.Write($"\n{new string(' ', 8)}{drawingInfo[0].ToString().Replace("{0}", $" style=\"position: absolute; left: {left}; top: {top}; width: {widthActual}; height: {heightActual};{(!string.IsNullOrEmpty(htmlStyleTransform) ? $" transform:{htmlStyleTransform};" : string.Empty)}{(isHidden ? " visibility: hidden;" : string.Empty)}\"")}");
            }
        }

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
            margin: 10px auto;
            width: fit-content;
            font-size: 20px;
            font-weight: bold;
            font-family: monospace;
            text-align: center;
            border-bottom: thick solid transparent;
        }

        table {
            width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
        }

        td {
            padding: 0;
            color: black;
            text-align: left;
            vertical-align: bottom;
            background-color: transparent;
            border: thin solid lightgray;
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
            this.BufferSize = 65536;
            this.ConvertStyles = true;
            this.ConvertSizes = true;
            this.ConvertPictures = true;
            this.ConvertSheetTitles = true;
            this.ConvertHiddenSheets = false;
            this.ConvertFirstSheetOnly = false;
            this.ConvertHtmlBodyOnly = false;
            this.RoundingDigits = 2;
        }

        #region Public Fields

        /// <summary>
        /// Gets or sets the Html page title.
        /// </summary>
        public string PageTitle { get; set; }

        /// <summary>
        /// Gets or sets the preset CSS styles of the Html.
        /// </summary>
        public string PresetStyles { get; set; }

        /// <summary>
        /// Gets or sets the error message that will be written to the Html if the conversion fails. Any instances of the text "{EXCEPTION}" will be replaced by the exception message.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets the encoding to use when writing the Html string.
        /// </summary>
        public System.Text.Encoding Encoding { get; set; }

        /// <summary>
        /// Gets or sets the buffer size to use when writing the Html string.
        /// </summary>
        public int BufferSize { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx styles to Html styles.
        /// </summary>
        public bool ConvertStyles { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx cell sizes to Html sizes.
        /// </summary>
        public bool ConvertSizes { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx pictures to Html images.
        /// </summary>
        public bool ConvertPictures { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx sheet names to Html titles.
        /// </summary>
        public bool ConvertSheetTitles { get; set; }

        /// <summary>
        /// Gets or sets whether to convert Xlsx hidden sheets.
        /// </summary>
        public bool ConvertHiddenSheets { get; set; }

        /// <summary>
        /// Gets or sets whether to only convert the first Xlsx sheet.
        /// </summary>
        public bool ConvertFirstSheetOnly { get; set; }

        /// <summary>
        /// Gets or sets whether to only convert to the Html body element.
        /// </summary>
        public bool ConvertHtmlBodyOnly { get; set; }

        /// <summary>
        /// Gets or sets the number of digits to round the numbers to, or to not use rounding if the value is negative.
        /// </summary>
        public int RoundingDigits { get; set; }

        /// <summary>
        /// Gets a new instance of <see cref="ConverterConfig"/> with default settings.
        /// </summary>
        public static ConverterConfig DefaultSettings { get { return new ConverterConfig(); } }

        #endregion
    }

    /// <summary>
    /// The progress callback event arguments of the Xlsx to Html converter.
    /// </summary>
    public class ConverterProgressCallbackEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConverterProgressCallbackEventArgs"/> class with specific numbers of current progress.
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
        /// Gets the current progress in percentage, ranging from 0 to 100.
        /// </summary>
        public double ProgressPercent
        {
            get
            {
                return Math.Max(0, Math.Min(100, (double)(CurrentSheet - 1) / TotalSheets * 100 + (double)CurrentRow / TotalRows * (100 / (double)TotalSheets)));
            }
        }

        /// <summary>
        /// Gets the 1-indexed number of the current sheet.
        /// </summary>
        public int CurrentSheet { get; }

        /// <summary>
        /// Gets the total amount of the sheets in the Xlsx file.
        /// </summary>
        public int TotalSheets { get; }

        /// <summary>
        /// Gets the 1-indexed number of the current row in the current sheet.
        /// </summary>
        public int CurrentRow { get; }

        /// <summary>
        /// Gets the total amount of the rows in the current sheet.
        /// </summary>
        public int TotalRows { get; }

        #endregion
    }
}
