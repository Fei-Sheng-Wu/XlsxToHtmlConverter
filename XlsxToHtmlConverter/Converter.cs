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
        /// Converts a local Xlsx file to Html string.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed at the cost of memory.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, bool loadIntoMemory = true)
        {
            ConvertXlsx(fileName, outputHtml, ConverterConfig.DefaultSettings, null, loadIntoMemory);
        }

        /// <summary>
        /// Converts a local Xlsx file to Html string with specific configuartions.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed at the cost of memory.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, ConverterConfig config, bool loadIntoMemory = true)
        {
            ConvertXlsx(fileName, outputHtml, config, null, loadIntoMemory);
        }

        /// <summary>
        /// Converts a local Xlsx file to Html string with progress callback event.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed at the cost of memory.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, EventHandler<ConverterProgressCallbackEventArgs> progressCallback, bool loadIntoMemory = true)
        {
            ConvertXlsx(fileName, outputHtml, ConverterConfig.DefaultSettings, progressCallback, loadIntoMemory);
        }

        /// <summary>
        /// Converts a local Xlsx file to Html string with specific configuartions and progress callback event.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed at the cost of memory.</param>
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
        /// Converts a stream Xlsx file to Html string.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml)
        {
            ConvertXlsx(inputXlsx, outputHtml, ConverterConfig.DefaultSettings, null);
        }

        /// <summary>
        /// Converts a stream Xlsx file to Html string with specific configurations.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml, ConverterConfig config)
        {
            ConvertXlsx(inputXlsx, outputHtml, config, null);
        }

        /// <summary>
        /// Converts a stream Xlsx file to Html string with progress callback event.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml, EventHandler<ConverterProgressCallbackEventArgs> progressCallback)
        {
            ConvertXlsx(inputXlsx, outputHtml, ConverterConfig.DefaultSettings, progressCallback);
        }

        /// <summary>
        /// Converts a stream Xlsx file to Html string with specific configurations and progress callback event.
        /// </summary>
        /// <param name="inputXlsx">The input stream of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        public static void ConvertXlsx(Stream inputXlsx, Stream outputHtml, ConverterConfig config, EventHandler<ConverterProgressCallbackEventArgs> progressCallback)
        {
            ConverterConfig configClone = config == null ? ConverterConfig.DefaultSettings : config.Clone();

            StreamWriter writer = new StreamWriter(outputHtml, configClone.Encoding, configClone.BufferSize);
            writer.BaseStream.Seek(0, SeekOrigin.Begin);
            writer.BaseStream.SetLength(0);

            try
            {
                writer.AutoFlush = true;
                writer.Write(!configClone.ConvertHtmlBodyOnly ? $@"<!DOCTYPE html>
<html>

<head>
    <meta charset=""UTF-8"">
    <title>{configClone.PageTitle}</title>

    <style>
        {configClone.PresetStyles}
    </style>
</head>
<body>" : $"<style>\n{configClone.PresetStyles}\n</style>");

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(inputXlsx, false))
                {
                    WorkbookPart workbook = document.WorkbookPart;
                    IEnumerable<Sheet> sheets = workbook.Workbook.Descendants<Sheet>();

                    DocumentFormat.OpenXml.Drawing.Color2Type[] themeColors = null;
                    if (workbook.ThemePart != null && workbook.ThemePart.Theme != null && workbook.ThemePart.Theme.ThemeElements != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme != null)
                    {
                        themeColors = new DocumentFormat.OpenXml.Drawing.Color2Type[12] {
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Light1Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Dark1Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Light2Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Dark2Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent1Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent2Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent3Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent4Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent5Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Accent6Color,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.Hyperlink,
                            workbook.ThemePart.Theme.ThemeElements.ColorScheme.FollowedHyperlinkColor
                        };
                    }

                    Stylesheet stylesheet = workbook.WorkbookStylesPart != null && workbook.WorkbookStylesPart.Stylesheet != null ? workbook.WorkbookStylesPart.Stylesheet : null;
                    (Dictionary<string, string>, string, uint)[] stylesheetCellFormats = new (Dictionary<string, string>, string, uint)[stylesheet != null && stylesheet.CellFormats != null && stylesheet.CellFormats.HasChildren ? stylesheet.CellFormats.ChildElements.Count : 0];
                    for (int stylesheetFormatIndex = 0; stylesheetFormatIndex < stylesheetCellFormats.Length; stylesheetFormatIndex++)
                    {
                        if (stylesheet.CellFormats.ChildElements[stylesheetFormatIndex] is CellFormat cellFormat)
                        {
                            Fill fill = (cellFormat.ApplyFill == null || (cellFormat.ApplyFill.HasValue && cellFormat.ApplyFill.Value)) && cellFormat.FillId != null && cellFormat.FillId.HasValue && stylesheet.Fills != null && stylesheet.Fills.HasChildren && cellFormat.FillId.Value < stylesheet.Fills.ChildElements.Count ? (Fill)stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value] : null;
                            Font font = (cellFormat.ApplyFont == null || (cellFormat.ApplyFont.HasValue && cellFormat.ApplyFont.Value)) && cellFormat.FontId != null && cellFormat.FontId.HasValue && stylesheet.Fonts != null && stylesheet.Fonts.HasChildren && cellFormat.FontId.Value < stylesheet.Fonts.ChildElements.Count ? (Font)stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value] : null;
                            Border border = (cellFormat.ApplyBorder == null || (cellFormat.ApplyBorder.HasValue && cellFormat.ApplyBorder.Value)) && cellFormat.BorderId != null && cellFormat.BorderId.HasValue && stylesheet.Borders != null && stylesheet.Borders.HasChildren && cellFormat.BorderId.Value < stylesheet.Borders.ChildElements.Count ? (Border)stylesheet.Borders.ChildElements[(int)cellFormat.BorderId.Value] : null;
                            stylesheetCellFormats[stylesheetFormatIndex] = (CellFormatToHtml(fill, font, border, cellFormat.ApplyAlignment == null || (cellFormat.ApplyAlignment.HasValue && cellFormat.ApplyAlignment.Value) ? cellFormat.Alignment : null, out string cellValueContainer, themeColors, configClone), cellValueContainer, cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.HasValue && (cellFormat.ApplyNumberFormat == null || (cellFormat.ApplyNumberFormat.HasValue && cellFormat.ApplyNumberFormat.Value)) ? cellFormat.NumberFormatId.Value : 0);
                        }
                    }
                    (Dictionary<string, string>, string)[] stylesheetDifferentialFormats = new (Dictionary<string, string>, string)[stylesheet != null && stylesheet.DifferentialFormats != null && stylesheet.DifferentialFormats.HasChildren ? stylesheet.DifferentialFormats.ChildElements.Count : 0];
                    for (int stylesheetDifferentialFormatIndex = 0; stylesheetDifferentialFormatIndex < stylesheetDifferentialFormats.Length; stylesheetDifferentialFormatIndex++)
                    {
                        if (stylesheet.DifferentialFormats.ChildElements[stylesheetDifferentialFormatIndex] is DifferentialFormat differentialFormat)
                        {
                            stylesheetDifferentialFormats[stylesheetDifferentialFormatIndex] = (CellFormatToHtml(differentialFormat.Fill, differentialFormat.Font, differentialFormat.Border, differentialFormat.Alignment, out string cellValueContainer, themeColors, configClone), cellValueContainer);
                        }
                    }
                    Dictionary<uint, string> stylesheetNumberingFormats = new Dictionary<uint, string>();
                    Dictionary<uint, string> stylesheetNumberingFormatsDateTime = new Dictionary<uint, string>();
                    Dictionary<string, (int, int, int, bool, int, int, int, int, bool, bool, List<string>)> stylesheetNumberingFormatsNumber = new Dictionary<string, (int, int, int, bool, int, int, int, int, bool, bool, List<string>)>();
                    if (configClone.ConvertNumberFormats && stylesheet != null && stylesheet.NumberingFormats != null)
                    {
                        foreach (NumberingFormat numberingFormat in stylesheet.NumberingFormats.Descendants<NumberingFormat>())
                        {
                            if (numberingFormat.NumberFormatId != null && numberingFormat.NumberFormatId.HasValue)
                            {
                                stylesheetNumberingFormats.Add(numberingFormat.NumberFormatId.Value, numberingFormat.FormatCode != null && numberingFormat.FormatCode.HasValue ? System.Web.HttpUtility.HtmlDecode(numberingFormat.FormatCode.Value) : string.Empty);
                            }
                        }
                    }

                    SharedStringTable sharedStringTable = workbook.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() is SharedStringTablePart sharedStringTablePart ? sharedStringTablePart.SharedStringTable : null;
                    (string, string)[] cellValueSharedStrings = new (string, string)[sharedStringTable != null && sharedStringTable.HasChildren ? sharedStringTable.ChildElements.Count : 0];
                    for (int sharedStringIndex = 0; sharedStringIndex < cellValueSharedStrings.Length; sharedStringIndex++)
                    {
                        if (sharedStringTable.ChildElements[sharedStringIndex] is SharedStringItem sharedString)
                        {
                            string cellValue = string.Empty;
                            string cellValueRaw = string.Empty;
                            if (configClone.ConvertStyles && sharedString.HasChildren)
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
                                        Dictionary<string, string> htmlStylesRun = new Dictionary<string, string>();
                                        string cellValueContainer = "{0}";
                                        if (configClone.ConvertStyles && run.RunProperties is RunProperties runProperties)
                                        {
                                            if (runProperties.GetFirstChild<RunFont>() is RunFont runFont && runFont.Val != null && runFont.Val.HasValue)
                                            {
                                                htmlStylesRun.Add("font-family", runFont.Val.Value);
                                            }
                                            htmlStylesRun = JoinHtmlAttributes(htmlStylesRun, FontToHtml(runProperties.GetFirstChild<Color>(), runProperties.GetFirstChild<FontSize>(), runProperties.GetFirstChild<Bold>(), runProperties.GetFirstChild<Italic>(), runProperties.GetFirstChild<Strike>(), runProperties.GetFirstChild<Underline>(), out cellValueContainer, themeColors, configClone));
                                        }

                                        string htmlStylesRunString = GetHtmlAttributesString(htmlStylesRun, false, -1);
                                        cellValue += $"<span{(!string.IsNullOrEmpty(htmlStylesRunString) ? $" style=\"{htmlStylesRunString}\"" : string.Empty)}>{cellValueContainer.Replace("{0}", GetEscapedString(run.Text.Text))}</span>";
                                        cellValueRaw += run.Text.Text;
                                        runLast = run;
                                    }
                                }
                            }
                            else
                            {
                                string text = sharedString.Text != null ? sharedString.Text.Text : string.Empty;
                                cellValue = GetEscapedString(text);
                                cellValueRaw = text;
                            }
                            cellValueSharedStrings[sharedStringIndex] = (cellValue, cellValueRaw);
                        }
                    }

                    int sheetIndex = 0;
                    int sheetsCount = configClone.ConvertFirstSheetOnly ? Math.Min(1, sheets.Count()) : sheets.Count();
                    foreach (Sheet sheet in sheets)
                    {
                        sheetIndex++;
                        if ((configClone.ConvertFirstSheetOnly && sheetIndex > 1) || sheet.Id == null || !sheet.Id.HasValue || !(workbook.GetPartById(sheet.Id.Value) is WorksheetPart worksheetPart) || (configClone.ConvertHiddenSheets && sheet.State != null && sheet.State.HasValue && sheet.State.Value != SheetStateValues.Visible))
                        {
                            continue;
                        }
                        string sheetName = sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : string.Empty;
                        string sheetTabColor = string.Empty;
                        int[] sheetDimension = null;

                        double columnWidthDefault = 64;
                        double rowHeightDefault = 20;
                        List<(int, int, double)> columns = new List<(int, int, double)>();
                        Dictionary<(int, int), int[]> mergeCells = new Dictionary<(int, int), int[]>();
                        List<(ConditionalFormatting, List<int[]>, IEnumerable<ConditionalFormattingRule>)> conditionalFormattings = new List<(ConditionalFormatting, List<int[]>, IEnumerable<ConditionalFormattingRule>)>();

                        Worksheet worksheet = worksheetPart.Worksheet;
                        foreach (OpenXmlElement worksheetDescendant in worksheet.Descendants())
                        {
                            if (worksheetDescendant is SheetDimension worksheetDimension && worksheetDimension.Reference != null && worksheetDimension.Reference.HasValue)
                            {
                                sheetDimension = new int[4];
                                GetReferenceRange(worksheetDimension.Reference.Value, out sheetDimension[0], out sheetDimension[1], out sheetDimension[2], out sheetDimension[3]);
                            }
                            else if (worksheetDescendant is SheetProperties worksheetProperties && configClone.ConvertSheetTitles && worksheetProperties.TabColor != null)
                            {
                                sheetTabColor = ColorTypeToHtml(worksheetProperties.TabColor, themeColors, configClone);
                            }
                            else if (worksheetDescendant is SheetFormatProperties worksheetFormatProperties)
                            {
                                columnWidthDefault = RoundNumber(worksheetFormatProperties.DefaultColumnWidth != null && worksheetFormatProperties.DefaultColumnWidth.HasValue ? worksheetFormatProperties.DefaultColumnWidth.Value : (worksheetFormatProperties.BaseColumnWidth != null && worksheetFormatProperties.BaseColumnWidth.HasValue ? worksheetFormatProperties.BaseColumnWidth.Value : columnWidthDefault), configClone.RoundingDigits);
                                rowHeightDefault = RoundNumber(worksheetFormatProperties.DefaultRowHeight != null && worksheetFormatProperties.DefaultRowHeight.HasValue ? worksheetFormatProperties.DefaultRowHeight.Value / 72 * 96 : rowHeightDefault, configClone.RoundingDigits);
                            }
                            else if (worksheetDescendant is Columns columnsGroup && config.ConvertSizes)
                            {
                                foreach (Column column in columnsGroup.Descendants<Column>())
                                {
                                    bool isHidden = (column.Collapsed != null && column.Collapsed.HasValue && column.Collapsed.Value) || (column.Hidden != null && column.Hidden.HasValue && column.Hidden.Value);
                                    if ((column.Width != null && column.Width.HasValue && (column.CustomWidth == null || (column.CustomWidth.HasValue && column.CustomWidth.Value))) || isHidden)
                                    {
                                        columns.Add((column.Min != null && column.Min.HasValue ? (int)column.Min.Value : int.MinValue, column.Max != null && column.Max.HasValue ? (int)column.Max.Value : int.MaxValue, isHidden ? 0 : column.Width.Value));
                                    }
                                }
                            }
                            else if (worksheetDescendant is MergeCells mergeCellsGroup)
                            {
                                foreach (MergeCell mergeCell in mergeCellsGroup.Descendants<MergeCell>())
                                {
                                    if (mergeCell.Reference == null || !mergeCell.Reference.HasValue)
                                    {
                                        continue;
                                    }

                                    GetReferenceRange(mergeCell.Reference.Value, out int mergeCellFromColumn, out int mergeCellFromRow, out int mergeCellToColumn, out int mergeCellToRow);
                                    mergeCells[(mergeCellFromColumn, mergeCellFromRow)] = new int[2] { mergeCellToColumn - mergeCellFromColumn + 1, mergeCellToRow - mergeCellFromRow + 1 };
                                    for (int i = mergeCellFromColumn; i <= mergeCellToColumn; i++)
                                    {
                                        for (int j = mergeCellFromRow; j <= mergeCellToRow; j++)
                                        {
                                            if (!mergeCells.ContainsKey((i, j)))
                                            {
                                                mergeCells.Add((i, j), null);
                                            }
                                        }
                                    }
                                }
                            }
                            else if (worksheetDescendant is ConditionalFormatting worksheetConditionalFormatting)
                            {
                                if (worksheetConditionalFormatting.SequenceOfReferences != null && worksheetConditionalFormatting.SequenceOfReferences.HasValue)
                                {
                                    List<int[]> sequence = new List<int[]>();
                                    foreach (string references in worksheetConditionalFormatting.SequenceOfReferences.Items)
                                    {
                                        int[] range = new int[4];
                                        GetReferenceRange(references, out range[0], out range[1], out range[2], out range[3]);
                                        sequence.Add(range);
                                    }
                                    conditionalFormattings.Add((worksheetConditionalFormatting, sequence, worksheetConditionalFormatting.Descendants<ConditionalFormattingRule>()));
                                }
                            }
                        }

                        Dictionary<int, double> drawingColumnMarkers = new Dictionary<int, double>();
                        Dictionary<int, double> drawingRowMarkers = new Dictionary<int, double>();
                        if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                        {
                            foreach (OpenXmlElement drawing in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants())
                            {
                                if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor && oneCellAnchor.FromMarker != null)
                                {
                                    if (oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) && !drawingColumnMarkers.ContainsKey(columnId))
                                    {
                                        drawingColumnMarkers.Add(columnId, double.NaN);
                                    }
                                    if (oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) && !drawingRowMarkers.ContainsKey(rowId))
                                    {
                                        drawingRowMarkers.Add(rowId, double.NaN);
                                    }
                                }
                                else if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor)
                                {
                                    if (twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) && !drawingColumnMarkers.ContainsKey(fromColumnId))
                                    {
                                        drawingColumnMarkers.Add(fromColumnId, double.NaN);
                                    }
                                    if (twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) && !drawingRowMarkers.ContainsKey(fromRowId))
                                    {
                                        drawingRowMarkers.Add(fromRowId, double.NaN);
                                    }
                                    if (twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) && !drawingColumnMarkers.ContainsKey(toColumnId))
                                    {
                                        drawingColumnMarkers.Add(toColumnId, double.NaN);
                                    }
                                    if (twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) && !drawingRowMarkers.ContainsKey(toRowId))
                                    {
                                        drawingRowMarkers.Add(toRowId, double.NaN);
                                    }
                                }
                            }
                        }

                        writer.Write($"\n{new string(' ', 4)}<h5{(!string.IsNullOrEmpty(sheetTabColor) ? $" style=\"border-bottom-color: {sheetTabColor};\"" : string.Empty)}>{(!string.IsNullOrEmpty(sheetName) ? sheetName : "Untitled Sheet")}</h5>");
                        writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                        writer.Write($"\n{new string(' ', 8)}<table>");

                        if (sheetDimension == null)
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

                        double[] columnWidths = new double[sheetDimension[2] - sheetDimension[0] + 1];
                        for (int columnWidthIndex = 0; columnWidthIndex < columnWidths.Length; columnWidthIndex++)
                        {
                            columnWidths[columnWidthIndex] = columnWidthDefault;
                        }
                        if (configClone.ConvertSizes)
                        {
                            foreach ((int, int, double) columnInfo in columns)
                            {
                                for (int i = Math.Max(sheetDimension[0], columnInfo.Item1); i <= Math.Min(sheetDimension[2], columnInfo.Item2); i++)
                                {
                                    columnWidths[i - sheetDimension[0]] = columnInfo.Item3;
                                }
                            }
                            double columnWidthsTotal = columnWidths.Sum();
                            double columbWidthsAccumulation = 0;
                            for (int columnWidthIndex = 0; columnWidthIndex < columnWidths.Length; columnWidthIndex++)
                            {
                                columnWidths[columnWidthIndex] = RoundNumber(columnWidths[columnWidthIndex] / columnWidthsTotal * 100, configClone.RoundingDigits);
                                columbWidthsAccumulation += columnWidths[columnWidthIndex];
                                if (drawingColumnMarkers.ContainsKey(sheetDimension[0] + columnWidthIndex))
                                {
                                    drawingColumnMarkers[sheetDimension[0] + columnWidthIndex] = RoundNumber(columbWidthsAccumulation, configClone.RoundingDigits);
                                }
                            }
                        }

                        int rowIndex = sheetDimension[1];
                        double rowHeightsAccumulation = 0;
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
                                    if (configClone.ConvertSizes)
                                    {
                                        rowHeightsAccumulation += rowHeightDefault + 0.8;
                                        if (drawingRowMarkers.ContainsKey(additionalRowIndex))
                                        {
                                            drawingRowMarkers[additionalRowIndex] = RoundNumber(rowHeightsAccumulation, configClone.RoundingDigits);
                                        }
                                    }

                                    writer.Write($"\n{new string(' ', 12)}<tr>");
                                    for (int additionalColumnIndex = 0; additionalColumnIndex < columnWidths.Length; additionalColumnIndex++)
                                    {
                                        writer.Write($"\n{new string(' ', 16)}<td style=\"height: {rowHeightDefault}px; width: {columnWidths[additionalColumnIndex]}%;\"></td>");
                                    }
                                    writer.Write($"\n{new string(' ', 12)}</tr>");
                                }
                                rowIndex = (int)row.RowIndex.Value;
                            }
                            double cellHeightActual = configClone.ConvertSizes ? RoundNumber((row.CustomHeight == null || (row.CustomHeight.HasValue && row.CustomHeight.Value)) && row.Height != null && row.Height.HasValue ? row.Height.Value / 72 * 96 : rowHeightDefault, configClone.RoundingDigits) : double.NaN;
                            if (configClone.ConvertSizes)
                            {
                                rowHeightsAccumulation += cellHeightActual + 0.8;
                                if (drawingRowMarkers.ContainsKey(rowIndex))
                                {
                                    drawingRowMarkers[rowIndex] = RoundNumber(rowHeightsAccumulation, configClone.RoundingDigits);
                                }
                            }

                            writer.Write($"\n{new string(' ', 12)}<tr>");

                            Cell[] cells = new Cell[columnWidths.Length];
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                if (cell.CellReference != null && cell.CellReference.HasValue)
                                {
                                    int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
                                    if (cellColumnIndex >= sheetDimension[0] && cellColumnIndex <= sheetDimension[2])
                                    {
                                        cells[cellColumnIndex - sheetDimension[0]] = cell;
                                    }
                                }
                            }
                            for (int cellIndex = sheetDimension[0]; cellIndex <= sheetDimension[2]; cellIndex++)
                            {
                                string cellColumnName = string.Empty;
                                int cellColumnIndex = cellIndex;
                                while (cellColumnIndex > 0)
                                {
                                    int modulo = (cellColumnIndex - 1) % 26;
                                    cellColumnName = (char)(65 + modulo) + cellColumnName;
                                    cellColumnIndex = (cellColumnIndex - modulo) / 26;
                                }
                                cells[cellIndex - sheetDimension[0]] = cells[cellIndex - sheetDimension[0]] ?? new Cell() { CellValue = new CellValue(string.Empty) };
                                cells[cellIndex - sheetDimension[0]].CellReference = $"{cellColumnName}{rowIndex}";
                            }

                            int columnIndex = sheetDimension[0];
                            foreach (Cell cell in cells)
                            {
                                columnIndex = GetColumnIndex(cell.CellReference.Value);
                                double cellWidthActual = configClone.ConvertSizes ? columnWidths[columnIndex - sheetDimension[0]] : double.NaN;

                                int columnSpanned = 1;
                                int rowSpanned = 1;
                                if (mergeCells.ContainsKey((columnIndex, rowIndex)))
                                {
                                    if (!(mergeCells[(columnIndex, rowIndex)] is int[] mergeCellInfo))
                                    {
                                        continue;
                                    }
                                    columnSpanned = mergeCellInfo[0];
                                    rowSpanned = mergeCellInfo[1];
                                }

                                int styleIndex = cell.StyleIndex != null && cell.StyleIndex.HasValue ? (int)cell.StyleIndex.Value : (row.StyleIndex != null && row.StyleIndex.HasValue ? (int)row.StyleIndex.Value : -1);

                                string numberFormatCode = string.Empty;
                                bool isNumberFormatDefaultDateTime = false;
                                if (configClone.ConvertNumberFormats && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length)
                                {
                                    switch (stylesheetCellFormats[styleIndex].Item3)
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
                                            numberFormatCode = "MM-dd-yy";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 15:
                                            numberFormatCode = "d-MMM-yy";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 16:
                                            numberFormatCode = "d-MMM";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 17:
                                            numberFormatCode = "MMM-yy";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 18:
                                            numberFormatCode = "h:mm AM/PM";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 19:
                                            numberFormatCode = "h:mm:ss AM/PM";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 20:
                                            numberFormatCode = "h:mm";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 21:
                                            numberFormatCode = "h:mm:ss";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 22:
                                            numberFormatCode = "M/d/yy h:mm";
                                            isNumberFormatDefaultDateTime = true;
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
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 46:
                                            numberFormatCode = "[h]:mm:ss";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 47:
                                            numberFormatCode = "mmss.0";
                                            isNumberFormatDefaultDateTime = true;
                                            break;
                                        case 48:
                                            numberFormatCode = "##0.0E+0";
                                            break;
                                        case 49:
                                            numberFormatCode = "@";
                                            break;
                                        default:
                                            if (stylesheetNumberingFormats.ContainsKey(stylesheetCellFormats[styleIndex].Item3))
                                            {
                                                numberFormatCode = stylesheetNumberingFormats[stylesheetCellFormats[styleIndex].Item3];
                                            }
                                            break;
                                    }
                                }

                                Dictionary<string, string> htmlStylesCell = new Dictionary<string, string>();
                                string cellValueContainer = "{0}";
                                if (configClone.ConvertStyles && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length)
                                {
                                    htmlStylesCell = !configClone.UseHtmlStyleClasses && stylesheetCellFormats[styleIndex].Item1 != null ? JoinHtmlAttributes(htmlStylesCell, stylesheetCellFormats[styleIndex].Item1) : htmlStylesCell;
                                    cellValueContainer = !string.IsNullOrEmpty(stylesheetCellFormats[styleIndex].Item2) ? cellValueContainer.Replace("{0}", stylesheetCellFormats[styleIndex].Item2) : cellValueContainer;
                                }

                                string cellValue = string.Empty;
                                string cellValueRaw = string.Empty;
                                if (cell.CellValue != null)
                                {
                                    cellValue = cell.CellValue.Text;

                                    bool isSharedString = false;
                                    if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.SharedString && int.TryParse(cellValue, out int sharedStringId) && sharedStringId >= 0 && sharedStringId < cellValueSharedStrings.Length)
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
                                        if ((isNumberFormatDefaultDateTime || (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Date)) && double.TryParse(cellValueRaw, out double cellValueDate))
                                        {
                                            if (!isNumberFormatDefaultDateTime && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length && stylesheetNumberingFormatsDateTime.ContainsKey(stylesheetCellFormats[styleIndex].Item3))
                                            {
                                                numberFormatCode = stylesheetNumberingFormatsDateTime[stylesheetCellFormats[styleIndex].Item3];
                                            }
                                            else if (!isNumberFormatDefaultDateTime)
                                            {
                                                int status = -1;
                                                string numberFormatCodeNew = string.Empty;
                                                for (int i = 0; i < numberFormatCode.Length; i++)
                                                {
                                                    if (numberFormatCode[i] != 'm')
                                                    {
                                                        status = -1;
                                                        numberFormatCodeNew += numberFormatCode[i];
                                                        continue;
                                                    }
                                                    else if (status < 0)
                                                    {
                                                        for (int j = i - 1; j >= 0; j--)
                                                        {
                                                            if (numberFormatCode[j] == 'h' || numberFormatCode[j] == 'd' || numberFormatCode[j] == 'y')
                                                            {
                                                                status = numberFormatCode[j] == 'h' ? 2 : 1;
                                                                break;
                                                            }
                                                        }
                                                        for (int j = status < 2 ? i + 1 : numberFormatCode.Length; j < numberFormatCode.Length; j++)
                                                        {
                                                            if (numberFormatCode[j] == 's' || numberFormatCode[j] == 'd' || numberFormatCode[j] == 'y')
                                                            {
                                                                status = numberFormatCode[j] == 's' ? 2 : 1;
                                                                break;
                                                            }
                                                        }
                                                        status = status < 0 ? 1 : status;
                                                    }
                                                    numberFormatCodeNew += status == 1 ? 'M' : 'm';
                                                }
                                                if (styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length)
                                                {
                                                    stylesheetNumberingFormatsDateTime.Add(stylesheetCellFormats[styleIndex].Item3, numberFormatCodeNew);
                                                }
                                                numberFormatCode = numberFormatCodeNew;
                                            }

                                            DateTime dateValue = DateTime.FromOADate(cellValueDate).Date;
                                            cellValue = GetEscapedString(dateValue.ToString(numberFormatCode));
                                        }
                                        else
                                        {
                                            bool isCellValueNumber = double.TryParse(cellValueRaw, out double cellValueNumber);
                                            string[] numberFormatCodeComponents = numberFormatCode.Split(';');
                                            if (numberFormatCodeComponents.Length > 1 && isCellValueNumber)
                                            {
                                                int indexComponent = cellValueNumber > 0 || (numberFormatCodeComponents.Length == 2 && cellValueNumber == 0) ? 0 : (cellValueNumber < 0 ? 1 : (numberFormatCodeComponents.Length > 2 ? 2 : -1));
                                                numberFormatCode = indexComponent >= 0 ? numberFormatCodeComponents[indexComponent] : numberFormatCode;
                                            }
                                            else
                                            {
                                                numberFormatCode = numberFormatCodeComponents.Length > 3 ? numberFormatCodeComponents[3] : numberFormatCode;
                                            }
                                            cellValue = GetEscapedString(GetFormattedNumber(cellValueRaw, numberFormatCode.Trim(), ref stylesheetNumberingFormatsNumber, out List<string> conditions));
                                            if (config.ConvertStyles && isCellValueNumber && conditions != null)
                                            {
                                                for (int i = 0; i < conditions.Count; i++)
                                                {
                                                    string conditionColor = string.Empty;
                                                    switch (conditions[i].ToLower())
                                                    {
                                                        case "black":
                                                            conditionColor = "rgb(0, 0, 0)";
                                                            break;
                                                        case "blue":
                                                            conditionColor = "rgb(0, 0, 255)";
                                                            break;
                                                        case "cyan":
                                                            conditionColor = "rgb(0, 255, 255)";
                                                            break;
                                                        case "green":
                                                            conditionColor = "rgb(0, 128, 0)";
                                                            break;
                                                        case "magenta":
                                                            conditionColor = "rgb(255, 0, 255)";
                                                            break;
                                                        case "red":
                                                            conditionColor = "rgb(255, 0, 0)";
                                                            break;
                                                        case "white":
                                                            conditionColor = "rgb(255, 255, 255)";
                                                            break;
                                                        case "yellow":
                                                            conditionColor = "rgb(255, 255, 0)";
                                                            break;
                                                    }
                                                    bool isConditionMet = true;
                                                    if (i + 1 < conditions.Count && !char.IsLetter(conditions[i + 1][0]))
                                                    {
                                                        i++;
                                                        if (conditions[i].StartsWith("="))
                                                        {
                                                            isConditionMet = double.TryParse(conditions[i].Substring(1, conditions[i].Length - 1), out double conditionValue) && cellValueNumber == conditionValue;
                                                        }
                                                        else if (conditions[i].StartsWith("<>"))
                                                        {
                                                            isConditionMet = double.TryParse(conditions[i].Substring(2, conditions[i].Length - 2), out double conditionValue) && cellValueNumber != conditionValue;
                                                        }
                                                        else if (conditions[i].StartsWith(">="))
                                                        {
                                                            isConditionMet = double.TryParse(conditions[i].Substring(2, conditions[i].Length - 2), out double conditionValue) && cellValueNumber >= conditionValue;
                                                        }
                                                        else if (conditions[i].StartsWith("<="))
                                                        {
                                                            isConditionMet = double.TryParse(conditions[i].Substring(2, conditions[i].Length - 2), out double conditionValue) && cellValueNumber <= conditionValue;
                                                        }
                                                        else if (conditions[i].StartsWith(">"))
                                                        {
                                                            isConditionMet = double.TryParse(conditions[i].Substring(1, conditions[i].Length - 1), out double conditionValue) && cellValueNumber > conditionValue;
                                                        }
                                                        else if (conditions[i].StartsWith("<"))
                                                        {
                                                            isConditionMet = double.TryParse(conditions[i].Substring(1, conditions[i].Length - 1), out double conditionValue) && cellValueNumber < conditionValue;
                                                        }
                                                    }
                                                    if (!string.IsNullOrEmpty(conditionColor) && isConditionMet)
                                                    {
                                                        htmlStylesCell["color"] = conditionColor;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else if (!isSharedString)
                                    {
                                        cellValue = GetEscapedString(cellValue);
                                    }
                                }

                                if (configClone.ConvertStyles)
                                {
                                    if (cell.DataType != null && cell.DataType.HasValue)
                                    {
                                        if (cell.DataType.Value == CellValues.Error || cell.DataType.Value == CellValues.Boolean)
                                        {
                                            htmlStylesCell.Add("text-align", "center");
                                        }
                                        else if (cell.DataType.Value == CellValues.Date || cell.DataType.Value == CellValues.Number)
                                        {
                                            htmlStylesCell.Add("text-align", "right");
                                        }
                                    }
                                    else if (isNumberFormatDefaultDateTime || double.TryParse(cellValueRaw, out double _))
                                    {
                                        htmlStylesCell.Add("text-align", "right");
                                    }

                                    int differentialStyleIndex = -1;
                                    foreach ((ConditionalFormatting, List<int[]>, IEnumerable<ConditionalFormattingRule>) conditionalFormatting in conditionalFormattings)
                                    {
                                        if (!conditionalFormatting.Item2.Any(x => columnIndex >= x[0] && columnIndex <= x[2] && rowIndex >= x[1] && rowIndex <= x[3]))
                                        {
                                            continue;
                                        }

                                        int priorityCurrent = int.MaxValue;
                                        foreach (ConditionalFormattingRule formattingRule in conditionalFormatting.Item3)
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
                                                if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.Equal)
                                                {
                                                    isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaEqual && cellValueRaw == formulaEqual.Text.Trim('"');
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.NotEqual)
                                                {
                                                    isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaNotEqual && cellValueRaw != formulaNotEqual.Text.Trim('"');
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.BeginsWith)
                                                {
                                                    isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaBeginsWith && cellValueRaw.StartsWith(formulaBeginsWith.Text.Trim('"'));
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.EndsWith)
                                                {
                                                    isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaEndsWith && cellValueRaw.EndsWith(formulaEndsWith.Text.Trim('"'));
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.ContainsText)
                                                {
                                                    isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaContainsText && cellValueRaw.Contains(formulaContainsText.Text.Trim('"'));
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.NotContains)
                                                {
                                                    isConditionMet = formattingRule.GetFirstChild<Formula>() is Formula formulaNotContains && !cellValueRaw.Contains(formulaNotContains.Text.Trim('"'));
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.GreaterThan)
                                                {
                                                    isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] > x[1]);
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.GreaterThanOrEqual)
                                                {
                                                    isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] >= x[1]);
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.LessThan)
                                                {
                                                    isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] < x[1]);
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.LessThanOrEqual)
                                                {
                                                    isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 1, x => x[0] <= x[1]);
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.Between)
                                                {
                                                    isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 2, x => x[0] >= Math.Min(x[1], x[2]) && x[0] <= Math.Max(x[1], x[2]));
                                                }
                                                else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.NotBetween)
                                                {
                                                    isConditionMet = GetNumberFormulaCondition(cellValueRaw, formattingRule.Descendants<Formula>(), 2, x => x[0] < Math.Min(x[1], x[2]) || x[0] > Math.Max(x[1], x[2]));
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
                                    if (differentialStyleIndex >= 0 && differentialStyleIndex < stylesheetDifferentialFormats.Length)
                                    {
                                        htmlStylesCell = JoinHtmlAttributes(htmlStylesCell, stylesheetDifferentialFormats[differentialStyleIndex].Item1);
                                        cellValueContainer = cellValueContainer.Replace("{0}", stylesheetDifferentialFormats[differentialStyleIndex].Item2);
                                    }
                                }

                                writer.Write($"\n{new string(' ', 16)}<td{(columnSpanned > 1 ? $" colspan=\"{columnSpanned}\"" : string.Empty)}{(rowSpanned > 1 ? $" rowspan=\"{rowSpanned}\"" : string.Empty)}{(configClone.UseHtmlStyleClasses && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length ? $" class=\"format-{styleIndex}\"" : string.Empty)} style=\"width: {(!double.IsNaN(cellWidthActual) && columnSpanned <= 1 ? $"{cellWidthActual}%" : "auto")}; height: {(!double.IsNaN(cellHeightActual) && rowSpanned <= 1 ? $"{cellHeightActual}px" : "auto")};{GetHtmlAttributesString(htmlStylesCell, true, -1)}\">{cellValueContainer.Replace("{0}", cellValue)}</td>");
                            }

                            writer.Write($"\n{new string(' ', 12)}</tr>");

                            progressCallback?.Invoke(document, new ConverterProgressCallbackEventArgs(sheetIndex, sheetsCount, rowIndex - sheetDimension[1] + 1, sheetDimension[3] - sheetDimension[1] + 1));
                        }

                        writer.Write($"\n{new string(' ', 8)}</table>");

                        if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                        {
                            foreach (OpenXmlElement drawing in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants())
                            {
                                if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absoluteAnchor)
                                {
                                    string left = absoluteAnchor.Position != null && absoluteAnchor.Position.X != null && absoluteAnchor.Position.X.HasValue ? $"{RoundNumber(absoluteAnchor.Position.X.Value / 914400.0 * 96, configClone.RoundingDigits)}px" : "0";
                                    string top = absoluteAnchor.Position != null && absoluteAnchor.Position.Y != null && absoluteAnchor.Position.Y.HasValue ? $"{RoundNumber(absoluteAnchor.Position.Y.Value / 914400.0 * 96, configClone.RoundingDigits)}px" : "0";
                                    string width = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cx != null && absoluteAnchor.Extent.Cx.HasValue ? $"{RoundNumber(absoluteAnchor.Extent.Cx.Value / 914400.0 * 96, configClone.RoundingDigits)}px" : "auto";
                                    string height = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cy != null && absoluteAnchor.Extent.Cy.HasValue ? $"{RoundNumber(absoluteAnchor.Extent.Cy.Value / 914400.0 * 96, configClone.RoundingDigits)}px" : "auto";
                                    DrawingsToHtml(worksheetPart, absoluteAnchor, writer, left, top, width, height, configClone);
                                }
                                else if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor)
                                {
                                    double left = configClone.ConvertSizes && oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) && drawingColumnMarkers.ContainsKey(columnId) ? drawingColumnMarkers[columnId] : double.NaN;
                                    double leftOffset = oneCellAnchor.FromMarker.ColumnOffset != null && int.TryParse(oneCellAnchor.FromMarker.ColumnOffset.Text, out int columnOffset) ? RoundNumber(columnOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                    double top = configClone.ConvertSizes && oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) && drawingRowMarkers.ContainsKey(rowId) ? drawingRowMarkers[rowId] : double.NaN;
                                    double topOffset = oneCellAnchor.FromMarker.RowOffset != null && int.TryParse(oneCellAnchor.FromMarker.RowOffset.Text, out int rowOffset) ? RoundNumber(rowOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                    string width = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cx != null && oneCellAnchor.Extent.Cx.HasValue ? $"{RoundNumber(oneCellAnchor.Extent.Cx.Value / 914400.0 * 96, configClone.RoundingDigits)}px" : "auto";
                                    string height = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cy != null && oneCellAnchor.Extent.Cy.HasValue ? $"{RoundNumber(oneCellAnchor.Extent.Cy.Value / 914400.0 * 96, configClone.RoundingDigits)}px" : "auto";
                                    DrawingsToHtml(worksheetPart, oneCellAnchor, writer, !double.IsNaN(left) ? $"calc({left}% + {leftOffset}px)" : "0", !double.IsNaN(top) ? $"{RoundNumber(top + topOffset, configClone.RoundingDigits)}px" : "0", width, height, configClone);
                                }
                                else if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor)
                                {
                                    double fromColumn = configClone.ConvertSizes && twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) && drawingColumnMarkers.ContainsKey(fromColumnId) ? drawingColumnMarkers[fromColumnId] : double.NaN;
                                    double fromColumnOffset = twoCellAnchor.FromMarker.ColumnOffset != null && int.TryParse(twoCellAnchor.FromMarker.ColumnOffset.Text, out int fromMarkerColumnOffset) ? RoundNumber(fromMarkerColumnOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                    double fromRow = configClone.ConvertSizes && twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) && drawingRowMarkers.ContainsKey(fromRowId) ? drawingRowMarkers[fromRowId] : double.NaN;
                                    double fromRowOffset = twoCellAnchor.FromMarker.RowOffset != null && int.TryParse(twoCellAnchor.FromMarker.RowOffset.Text, out int fromMarkerRowOffset) ? RoundNumber(fromMarkerRowOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                    double toColumn = configClone.ConvertSizes && twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) && drawingColumnMarkers.ContainsKey(toColumnId) ? drawingColumnMarkers[toColumnId] : double.NaN;
                                    double toColumnOffset = twoCellAnchor.ToMarker.ColumnOffset != null && int.TryParse(twoCellAnchor.ToMarker.ColumnOffset.Text, out int toMarkerColumnOffset) ? RoundNumber(toMarkerColumnOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                    double toRow = configClone.ConvertSizes && twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) && drawingRowMarkers.ContainsKey(toRowId) ? drawingRowMarkers[toRowId] : double.NaN;
                                    double toRowOffset = twoCellAnchor.ToMarker.RowOffset != null && int.TryParse(twoCellAnchor.ToMarker.RowOffset.Text, out int toMarkerRowOffset) ? RoundNumber(toMarkerRowOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                    DrawingsToHtml(worksheetPart, twoCellAnchor, writer, !double.IsNaN(fromColumn) ? $"calc({fromColumn}% + {fromColumnOffset}px)" : "0", !double.IsNaN(fromRow) ? $"{RoundNumber(fromRow + fromRowOffset, configClone.RoundingDigits)}px" : "0", !double.IsNaN(fromColumn) && !double.IsNaN(toColumn) ? $"calc({RoundNumber(toColumn - fromColumn, configClone.RoundingDigits)}% + {RoundNumber(toColumnOffset - fromColumnOffset, configClone.RoundingDigits)}px)" : "auto", !double.IsNaN(fromRow) && !double.IsNaN(toRow) ? $"{RoundNumber(toRow + toRowOffset - fromRow - fromRowOffset, configClone.RoundingDigits)}px" : "auto", configClone);
                                }
                            }
                        }

                        writer.Write($"\n{new string(' ', 4)}</div>");
                    }

                    if (configClone.UseHtmlStyleClasses)
                    {
                        writer.Write($"\n{new string(' ', 4)}<style>");
                        for (int stylesheetFormatIndex = 0; stylesheetFormatIndex < stylesheetCellFormats.Length; stylesheetFormatIndex++)
                        {
                            if (stylesheetCellFormats[stylesheetFormatIndex].Item1 != null)
                            {
                                writer.Write($"\n{new string(' ', 8)}.format-{stylesheetFormatIndex} {{");
                                writer.Write($"\n{new string(' ', 12)}{GetHtmlAttributesString(stylesheetCellFormats[stylesheetFormatIndex].Item1, false, 12)}");
                                writer.Write($"\n{new string(' ', 8)}}}{(stylesheetFormatIndex < stylesheetCellFormats.Length - 1 ? $"\n{new string(' ', 8)}" : string.Empty)}");
                            }
                        }
                        writer.Write($"\n{new string(' ', 4)}</style>");
                    }
                }

                writer.Write(!configClone.ConvertHtmlBodyOnly ? "\n</body>\n</html>" : string.Empty);
            }
            catch (Exception ex)
            {
                writer.BaseStream.Seek(0, SeekOrigin.Begin);
                writer.BaseStream.SetLength(0);
                writer.Write(configClone.ErrorMessage.Replace("{EXCEPTION}", ex.Message));
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

        private static double RoundNumber(double number, int digits)
        {
            return digits < 0 ? number : Math.Round(number, digits);
        }

        private static string GetEscapedString(string value)
        {
            return System.Web.HttpUtility.HtmlEncode(value).Replace(" ", "&nbsp;");
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
            return Math.Max(1, index + 1);
        }

        private static int GetRowIndex(string cell)
        {
            Match match = regexNumbers.Match(cell);
            return match.Success && int.TryParse(match.Value, out int index) ? index : 1;
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

        private static Dictionary<string, string> JoinHtmlAttributes(Dictionary<string, string> original, Dictionary<string, string> joining)
        {
            if (joining == null)
            {
                return original;
            }

            foreach (KeyValuePair<string, string> pair in joining)
            {
                original[pair.Key] = pair.Value;
            }
            return original;
        }

        private static string GetHtmlAttributesString(Dictionary<string, string> attributes, bool isAdditional, int indent)
        {
            if (attributes == null)
            {
                return string.Empty;
            }

            string htmlAttributes = string.Empty;
            foreach (KeyValuePair<string, string> pair in attributes)
            {
                htmlAttributes += $"{pair.Key}: {pair.Value};{(indent >= 0 ? $"\n{new string(' ', indent)}" : " ")}";
            }
            return isAdditional ? $" {htmlAttributes.TrimEnd()}" : htmlAttributes.TrimEnd();
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

        private static string GetFormattedNumber(string value, string format, ref Dictionary<string, (int, int, int, bool, int, int, int, int, bool, bool, List<string>)> formatsCalculated, out List<string> conditions)
        {
            if (string.IsNullOrEmpty(format))
            {
                conditions = null;
                return value;
            }

            bool isValueNumber = double.TryParse(value, out double valueNumber);
            if (!isValueNumber && !format.Contains("@"))
            {
                conditions = null;
                return value;
            }

            int infoValue = value.Length;
            Action actionUpdateValue = () =>
            {
                value = isValueNumber ? valueNumber.ToString(stringFormatNumber) : value;
                infoValue = value.IndexOf('.');
                infoValue = infoValue < 0 ? value.Length : infoValue;
            };
            actionUpdateValue.Invoke();

            bool isFormatCalculated = formatsCalculated.ContainsKey(format);
            (int, int, int, bool, int, int, int, int, bool, bool, List<string>) infoFormat = isFormatCalculated ? formatsCalculated[format] : (format.Length, format.Length, format.Length, false, -1, -1, -1, -1, false, false, new List<string>());

            (bool, string) valueScientific = (true, "0");
            bool isFormattingScientific = false;

            (string, string) valueFraction = (string.Empty, string.Empty);
            bool isFormattingFraction = false;

            int indexValue = 0;
            int indexFormat = 0;
            bool isIncreasing = true;
            Action actionUpdateInfo = () =>
            {
                if (isValueNumber && infoFormat.Item4)
                {
                    valueNumber *= 100;
                    actionUpdateValue.Invoke();
                }
                if (isValueNumber && infoFormat.Item5 >= 0)
                {
                    if (infoValue > 1)
                    {
                        valueScientific = (true, (infoValue - 1).ToString());
                        valueNumber /= Math.Pow(10, infoValue - 1);
                        actionUpdateValue.Invoke();
                    }
                    else if (infoValue > 0 && value.Length > infoValue && value[0] == '0')
                    {
                        int digit = 0;
                        for (int i = infoValue + 1; i < value.Length; i++)
                        {
                            if (value[i] != '0')
                            {
                                digit = i;
                                break;
                            }
                        }
                        if (digit > infoValue)
                        {
                            valueScientific = (false, (digit - infoValue).ToString());
                            valueNumber *= Math.Pow(10, digit - infoValue);
                            actionUpdateValue.Invoke();
                        }
                    }
                }
                if (isValueNumber && infoFormat.Item7 >= 0)
                {
                    double valueAbsolute = Math.Abs(valueNumber);
                    int valueFloor = (int)Math.Floor(valueAbsolute);
                    if (infoFormat.Item10)
                    {
                        double valueInteger = valueNumber >= 0 ? valueFloor : -valueFloor;
                        valueNumber -= valueNumber >= 0 ? valueFloor : -valueFloor;
                        valueAbsolute = Math.Abs(valueNumber);
                        valueFloor = (int)Math.Floor(valueAbsolute);
                        valueNumber = valueInteger;
                    }
                    else
                    {
                        valueNumber = 0;
                    }
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
                        fractionNumerator = valueNumber >= 0 ? valueFloor : -valueFloor;
                        fractionDenominator = 1;
                    }
                    else if (1 - maxError < valueAbsolute)
                    {
                        fractionNumerator = valueNumber >= 0 ? (valueFloor + 1) : -(valueFloor + 1);
                        fractionDenominator = 1;
                    }
                    else
                    {
                        int[] fractionParts = new int[4] { 0, 1, 1, 1 };
                        Action<int, int, int, int, Func<int, int, bool>> actionFindNewValue = (indexNumerator, indexDenominator, incrementNumerator, incrementDenominator, actionEvaluation) =>
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
                                fractionNumerator = valueNumber >= 0 ? valueFloor * middleDenominator + middleNumerator : -(valueFloor * middleDenominator + middleNumerator);
                                fractionDenominator = middleDenominator;
                                break;
                            }
                        }
                    }
                    valueFraction = (fractionNumerator.ToString(), fractionDenominator.ToString());
                    actionUpdateValue.Invoke();
                }

                indexValue = infoValue;
                indexFormat = infoFormat.Item2 - 1;
                isIncreasing = false;
            };
            if (isFormatCalculated)
            {
                actionUpdateInfo.Invoke();
            }

            string result = string.Empty;
            string resultFormatted = string.Empty;
            while (indexFormat < format.Length || !isFormatCalculated)
            {
                if (isFormatCalculated && !isIncreasing && isFormattingScientific && infoFormat.Item5 >= 0 && indexFormat > 0 && format[indexFormat - 1] == 'E' && (format[indexFormat] == '+' || format[indexFormat] == '-'))
                {
                    result = resultFormatted + new string(result.Reverse().ToArray());
                    resultFormatted = string.Empty;
                    indexFormat = infoFormat.Item5 + 1;
                    isIncreasing = true;
                    isFormattingScientific = false;
                    continue;
                }
                else if (isFormatCalculated && !isIncreasing && isFormattingFraction && infoFormat.Item7 >= 0 && indexFormat < infoFormat.Item6)
                {
                    result = resultFormatted + new string(result.Reverse().ToArray());
                    resultFormatted = string.Empty;
                    indexValue = -1;
                    indexFormat = infoFormat.Item7 + 1;
                    isIncreasing = true;
                    continue;
                }
                else if (indexFormat >= format.Length && !isFormatCalculated)
                {
                    infoFormat.Item2 = Math.Min(infoFormat.Item2, infoFormat.Item3 + 1);
                    infoFormat.Item5 = Math.Min(infoFormat.Item5, infoFormat.Item3);
                    infoFormat.Item8 = Math.Min(infoFormat.Item8, infoFormat.Item3);
                    infoFormat.Item11 = infoFormat.Item11.Count > 0 ? infoFormat.Item11 : null;
                    formatsCalculated.Add(format, infoFormat);
                    isFormatCalculated = true;
                    actionUpdateInfo.Invoke();
                    continue;
                }
                else if (indexFormat < 0)
                {
                    result = new string(result.Reverse().ToArray());
                    indexValue = infoValue;
                    indexFormat = infoFormat.Item2;
                    isIncreasing = true;
                    continue;
                }

                char formatChar = format[indexFormat];
                if ((isIncreasing && indexFormat + 1 < format.Length && formatChar == '\\') || (!isIncreasing && indexFormat > 0 && format[indexFormat - 1] == '\\'))
                {
                    result += isFormatCalculated ? format[isIncreasing ? indexFormat + 1 : indexFormat].ToString() : string.Empty;
                    indexFormat += isIncreasing ? 2 : -2;
                    continue;
                }
                else if (isIncreasing ? formatChar == '[' && indexFormat + 1 < format.Length : formatChar == ']' && indexFormat > 0)
                {
                    if (!isFormatCalculated)
                    {
                        infoFormat.Item11.Add(string.Empty);
                    }
                    do
                    {
                        indexFormat += isIncreasing ? 1 : -1;
                        if (!isFormatCalculated)
                        {
                            infoFormat.Item11[infoFormat.Item11.Count - 1] += format[indexFormat].ToString();
                        }
                    } while (isIncreasing ? indexFormat + 1 < format.Length && format[indexFormat + 1] != ']' : indexFormat > 0 && format[indexFormat - 1] != '[');
                    indexFormat += isIncreasing ? 2 : -2;
                    if (!isFormatCalculated)
                    {
                        infoFormat.Item11[infoFormat.Item11.Count - 1] = infoFormat.Item11[infoFormat.Item11.Count - 1].Trim();
                    }
                    continue;
                }
                else if (formatChar == '\"' && (isIncreasing ? indexFormat + 1 < format.Length : indexFormat > 0))
                {
                    do
                    {
                        indexFormat += isIncreasing ? 1 : -1;
                        result += isFormatCalculated ? format[indexFormat].ToString() : string.Empty;
                    }
                    while (isIncreasing ? indexFormat + 1 < format.Length && format[indexFormat + 1] != '\"' : indexFormat > 0 && format[indexFormat - 1] != '\"');
                    indexFormat += isIncreasing ? 2 : -2;
                    continue;
                }
                else if ((isIncreasing && indexFormat + 1 < format.Length && formatChar == '*') || (!isIncreasing && indexFormat > 0 && format[indexFormat - 1] == '*'))
                {
                    result += isFormatCalculated ? format[isIncreasing ? indexFormat + 1 : indexFormat].ToString() : string.Empty;
                    indexFormat += isIncreasing ? 2 : -2;
                    continue;
                }
                else if ((isIncreasing && indexFormat + 1 < format.Length && formatChar == '_') || (!isIncreasing && indexFormat > 0 && format[indexFormat - 1] == '_'))
                {
                    result += isFormatCalculated ? " " : string.Empty;
                    indexFormat += isIncreasing ? 2 : -2;
                    continue;
                }

                if (!isFormatCalculated)
                {
                    if (formatChar == '.')
                    {
                        infoFormat.Item2 = Math.Min(infoFormat.Item2, indexFormat);
                    }
                    else if (formatChar == '0' || formatChar == '#' || formatChar == '?')
                    {
                        infoFormat.Item1 = Math.Min(infoFormat.Item1, indexFormat);
                        infoFormat.Item3 = indexFormat;
                        infoFormat.Item6 = infoFormat.Item6 < 0 ? indexFormat : infoFormat.Item6;
                        infoFormat.Item9 = (indexFormat > infoFormat.Item2 && (formatChar == '0' || formatChar == '?') && infoFormat.Item5 < 0) || infoFormat.Item9;
                    }
                    else if (formatChar == ' ')
                    {
                        infoFormat.Item10 = indexFormat > infoFormat.Item1 || infoFormat.Item10;
                        if (infoFormat.Item10)
                        {
                            infoFormat.Item5 = infoFormat.Item5 < format.Length ? infoFormat.Item5 : infoFormat.Item3;
                            infoFormat.Item6 = infoFormat.Item7 < 0 ? Math.Max(infoFormat.Item6, indexFormat + 1) : infoFormat.Item6;
                            infoFormat.Item8 = infoFormat.Item8 < format.Length ? infoFormat.Item8 : infoFormat.Item3;
                        }
                    }
                    else if (formatChar == '%')
                    {
                        infoFormat.Item4 = true;
                    }
                    else if (formatChar == 'E' && isIncreasing && indexFormat + 1 < format.Length && (format[indexFormat + 1] == '+' || format[indexFormat + 1] == '-'))
                    {
                        infoFormat.Item5 = format.Length;
                        infoFormat.Item2 = Math.Min(infoFormat.Item2, indexFormat);
                        indexFormat++;
                    }
                    else if (formatChar == '/' && isIncreasing)
                    {
                        infoFormat.Item2 = Math.Min(infoFormat.Item2, infoFormat.Item6);
                        infoFormat.Item7 = indexFormat - 1;
                        infoFormat.Item8 = format.Length;
                    }
                }
                else
                {
                    if (formatChar == '@')
                    {
                        result += isIncreasing ? value : new string(value.Reverse().ToArray());
                    }
                    else if (formatChar == ';')
                    {
                        result += infoFormat.Item11 == null || infoFormat.Item11.Count <= 0 ? ";" : string.Empty;
                    }
                    else if (isValueNumber && formatChar == '.')
                    {
                        result += infoFormat.Item9 || (isIncreasing && indexValue + 1 < value.Length) ? "." : string.Empty;
                    }
                    else if (isValueNumber && formatChar == ',')
                    {
                        result += (isIncreasing ? indexValue + 1 < value.Length : indexValue > 0) ? "," : string.Empty;
                    }
                    else if (isValueNumber && formatChar == 'E' && isIncreasing && !isFormattingScientific && infoFormat.Item5 >= 0 && indexFormat + 1 < format.Length && (format[indexFormat + 1] == '+' || format[indexFormat + 1] == '-'))
                    {
                        resultFormatted = result + (valueScientific.Item1 ? (format[indexFormat + 1] == '-' ? "E" : "E+") : "E-");
                        result = string.Empty;
                        indexValue = valueScientific.Item2.Length;
                        indexFormat = infoFormat.Item5;
                        isIncreasing = false;
                        isFormattingScientific = true;
                        continue;
                    }
                    else if (isValueNumber && isIncreasing && !isFormattingFraction && infoFormat.Item7 >= 0 && indexFormat >= infoFormat.Item6 && indexFormat <= infoFormat.Item7)
                    {
                        resultFormatted = result;
                        result = string.Empty;
                        indexValue = valueFraction.Item1.Length;
                        indexFormat = infoFormat.Item7;
                        isIncreasing = false;
                        isFormattingFraction = true;
                        continue;
                    }
                    else if (isValueNumber && (formatChar == '0' || formatChar == '#' || formatChar == '?'))
                    {
                        indexValue += isIncreasing ? 1 : -1;
                        if (indexValue >= 0 && indexValue < (!isFormattingScientific ? (!isFormattingFraction ? value : (indexFormat > infoFormat.Item7 ? valueFraction.Item2 : valueFraction.Item1)) : valueScientific.Item2).Length && (isFormattingScientific || isFormattingFraction || formatChar == '0' || indexValue > 0 || value[indexValue] != '0' || infoFormat.Item9))
                        {
                            if (isIncreasing && !isFormattingFraction && (indexFormat >= infoFormat.Item3 || (isFormattingScientific && indexFormat + 2 < format.Length && format[indexFormat + 1] == 'E' && (format[indexFormat + 2] == '+' || format[indexFormat + 2] == '-'))) && indexValue + 1 < value.Length && int.TryParse(value[indexValue + 1].ToString(), out int next) && next > 4)
                            {
                                return GetFormattedNumber((valueNumber + (10 - next) / Math.Pow(10, indexValue + 1 - infoValue)).ToString(stringFormatNumber), format, ref formatsCalculated, out conditions);
                            }

                            result += (!isFormattingScientific ? (!isFormattingFraction ? value : (indexFormat > infoFormat.Item7 ? valueFraction.Item2 : valueFraction.Item1)) : valueScientific.Item2)[indexValue].ToString();
                            if (!isIncreasing && (!isFormattingScientific ? (!isFormattingFraction ? indexFormat <= infoFormat.Item1 : indexFormat <= infoFormat.Item6) : indexFormat - 2 >= 0 && format[indexFormat - 2] == 'E' && (format[indexFormat - 1] == '+' || format[indexFormat - 1] == '-')))
                            {
                                result += new string((!isFormattingScientific ? (!isFormattingFraction ? value : valueFraction.Item1) : valueScientific.Item2).Substring(0, indexValue).Reverse().ToArray());
                            }
                            else if (isIncreasing && isFormattingFraction && indexFormat >= infoFormat.Item8 && indexValue + 1 < valueFraction.Item2.Length)
                            {
                                result += valueFraction.Item2.Substring(indexValue + 1, valueFraction.Item2.Length - indexValue - 1);
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
                indexFormat += isIncreasing ? 1 : -1;
            }
            conditions = infoFormat.Item11;
            return result;
        }

        private static Dictionary<string, string> CellFormatToHtml(Fill fill, Font font, Border border, Alignment alignment, out string cellValueContainer, DocumentFormat.OpenXml.Drawing.Color2Type[] themeColors, ConverterConfig config)
        {
            Dictionary<string, string> htmlStyles = new Dictionary<string, string>();
            cellValueContainer = "{0}";
            if (fill != null && fill.PatternFill != null && (fill.PatternFill.PatternType == null || (fill.PatternFill.PatternType.HasValue && fill.PatternFill.PatternType.Value != PatternValues.None)))
            {
                string background = string.Empty;
                if (fill.PatternFill.ForegroundColor != null)
                {
                    background = ColorTypeToHtml(fill.PatternFill.ForegroundColor, themeColors, config);
                }
                if (string.IsNullOrEmpty(background) && fill.PatternFill.BackgroundColor != null)
                {
                    background = ColorTypeToHtml(fill.PatternFill.BackgroundColor, themeColors, config);
                }
                if (!string.IsNullOrEmpty(background))
                {
                    htmlStyles.Add("background-color", background);
                }
            }
            if (font != null)
            {
                if (font.FontName != null && font.FontName.Val != null && font.FontName.Val.HasValue)
                {
                    htmlStyles.Add("font-family", font.FontName.Val.Value);
                }
                htmlStyles = JoinHtmlAttributes(htmlStyles, FontToHtml(font.Color, font.FontSize, font.Bold, font.Italic, font.Strike, font.Underline, out cellValueContainer, themeColors, config));
            }
            if (border != null)
            {
                if (border.LeftBorder != null)
                {
                    string borderLeft = BorderPropertiesToHtml(border.LeftBorder, themeColors, config);
                    if (!string.IsNullOrEmpty(borderLeft))
                    {
                        htmlStyles.Add("border-left", borderLeft);
                    }
                }
                if (border.RightBorder != null)
                {
                    string borderRight = BorderPropertiesToHtml(border.RightBorder, themeColors, config);
                    if (!string.IsNullOrEmpty(borderRight))
                    {
                        htmlStyles.Add("border-right", borderRight);
                    }
                }
                if (border.TopBorder != null)
                {
                    string borderTop = BorderPropertiesToHtml(border.TopBorder, themeColors, config);
                    if (!string.IsNullOrEmpty(borderTop))
                    {
                        htmlStyles.Add("border-top", borderTop);
                    }
                }
                if (border.BottomBorder != null)
                {
                    string borderBottom = BorderPropertiesToHtml(border.BottomBorder, themeColors, config);
                    if (!string.IsNullOrEmpty(borderBottom))
                    {
                        htmlStyles.Add("border-bottom", borderBottom);
                    }
                }
            }
            if (alignment != null)
            {
                if (alignment.Horizontal != null && alignment.Horizontal.HasValue && alignment.Horizontal.Value != HorizontalAlignmentValues.General)
                {
                    htmlStyles.Add("text-align", alignment.Horizontal.Value == HorizontalAlignmentValues.Left ? "left" : (alignment.Horizontal.Value == HorizontalAlignmentValues.Right ? "right" : (alignment.Horizontal.Value == HorizontalAlignmentValues.Justify ? "justify" : "center")));
                }
                if (alignment.Vertical != null && alignment.Vertical.HasValue)
                {
                    htmlStyles.Add("vertical-align", alignment.Vertical.Value == VerticalAlignmentValues.Bottom ? "bottom" : (alignment.Vertical.Value == VerticalAlignmentValues.Top ? "top" : "middle"));
                }
                if (alignment.WrapText != null && alignment.WrapText.HasValue && alignment.WrapText.Value)
                {
                    htmlStyles.Add("word-wrap", "break-word");
                    htmlStyles.Add("white-space", "normal");
                }
                if (alignment.TextRotation != null && alignment.TextRotation.HasValue)
                {
                    cellValueContainer = cellValueContainer.Replace("{0}", $"<div style=\"width: fit-content; transform: rotate(-{RoundNumber(alignment.TextRotation.Value, config.RoundingDigits)}deg);\">{{0}}</div>");
                }
            }
            return htmlStyles;
        }

        private static Dictionary<string, string> FontToHtml(ColorType color, FontSize fontSize, Bold bold, Italic italic, Strike strike, Underline underline, out string cellValueContainer, DocumentFormat.OpenXml.Drawing.Color2Type[] themeColors, ConverterConfig config)
        {
            Dictionary<string, string> htmlStyles = new Dictionary<string, string>();
            cellValueContainer = "{0}";
            if (color != null)
            {
                string htmlColor = ColorTypeToHtml(color, themeColors, config);
                if (!string.IsNullOrEmpty(htmlColor))
                {
                    htmlStyles.Add("color", htmlColor);
                }
            }
            if (fontSize != null && fontSize.Val != null && fontSize.Val.HasValue)
            {
                htmlStyles.Add("font-size", $"{RoundNumber(fontSize.Val.Value / 72 * 96, config.RoundingDigits)}px");
            }
            if (bold != null)
            {
                htmlStyles.Add("font-weight", bold.Val == null || (bold.Val.HasValue && bold.Val.Value) ? "bold" : "normal");
            }
            if (italic != null)
            {
                htmlStyles.Add("font-style", italic.Val == null || (italic.Val.HasValue && italic.Val.Value) ? "italic" : "normal");
            }
            string htmlStylesTextDecoraion = string.Empty;
            if (strike != null)
            {
                htmlStylesTextDecoraion += strike.Val == null || (strike.Val.HasValue && strike.Val.Value) ? " line-through" : " none";
            }
            if (underline != null && underline.Val != null && underline.Val.HasValue)
            {
                if (underline.Val.Value == UnderlineValues.Double || underline.Val.Value == UnderlineValues.DoubleAccounting)
                {
                    cellValueContainer = cellValueContainer.Replace("{0}", $"<div style=\"width: fit-content; text-decoration: underline double;\">{{0}}</div>");
                }
                else if (underline.Val.Value != UnderlineValues.None)
                {
                    htmlStylesTextDecoraion += " underline";
                }
            }
            if (!string.IsNullOrEmpty(htmlStylesTextDecoraion))
            {
                htmlStyles.Add("text-decoration", htmlStylesTextDecoraion.TrimStart());
            }
            return htmlStyles;
        }

        private static string BorderPropertiesToHtml(BorderPropertiesType border, DocumentFormat.OpenXml.Drawing.Color2Type[] themeColors, ConverterConfig config)
        {
            if (border == null)
            {
                return string.Empty;
            }

            string htmlBorder = string.Empty;
            if (border.Style != null && border.Style.HasValue)
            {
                if (border.Style.Value == BorderStyleValues.Thick)
                {
                    htmlBorder += " thick solid";
                }
                else if (border.Style.Value == BorderStyleValues.Medium)
                {
                    htmlBorder += " medium solid";
                }
                else if (border.Style.Value == BorderStyleValues.MediumDashed || border.Style.Value == BorderStyleValues.MediumDashDot)
                {
                    htmlBorder += " medium dashed";
                }
                else if (border.Style.Value == BorderStyleValues.MediumDashDotDot)
                {
                    htmlBorder += " medium dotted";
                }
                else if (border.Style.Value == BorderStyleValues.Thin)
                {
                    htmlBorder += " thin solid";
                }
                else if (border.Style.Value == BorderStyleValues.Dashed || border.Style.Value == BorderStyleValues.DashDot || border.Style.Value == BorderStyleValues.SlantDashDot)
                {
                    htmlBorder += " thin dashed";
                }
                else if (border.Style.Value == BorderStyleValues.DashDotDot || border.Style.Value == BorderStyleValues.Hair)
                {
                    htmlBorder += " thin dotted";
                }
                else if (border.Style.Value == BorderStyleValues.Double)
                {
                    htmlBorder += " double";
                }
            }
            if (border.Color != null)
            {
                string htmlColor = ColorTypeToHtml(border.Color, themeColors, config);
                if (!string.IsNullOrEmpty(htmlColor))
                {
                    htmlBorder += $" {htmlColor}";
                }
            }
            return htmlBorder.TrimStart();
        }

        private static string ColorTypeToHtml(ColorType color, DocumentFormat.OpenXml.Drawing.Color2Type[] themeColors, ConverterConfig config)
        {
            if (color == null)
            {
                return string.Empty;
            }

            double[] result = new double[4] { 0, 0, 0, 1 };
            Action<double, double, double> actionUpdateColor = (red, green, blue) =>
            {
                result[0] = red;
                result[1] = green;
                result[2] = blue;
            };

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
                        actionUpdateColor.Invoke(0, 0, 0);
                        break;
                    case 1:
                        actionUpdateColor.Invoke(255, 255, 255);
                        break;
                    case 2:
                        actionUpdateColor.Invoke(255, 0, 0);
                        break;
                    case 3:
                        actionUpdateColor.Invoke(0, 255, 0);
                        break;
                    case 4:
                        actionUpdateColor.Invoke(0, 0, 255);
                        break;
                    case 5:
                        actionUpdateColor.Invoke(255, 255, 0);
                        break;
                    case 6:
                        actionUpdateColor.Invoke(255, 0, 255);
                        break;
                    case 7:
                        actionUpdateColor.Invoke(0, 255, 255);
                        break;
                    case 8:
                        actionUpdateColor.Invoke(0, 0, 0);
                        break;
                    case 9:
                        actionUpdateColor.Invoke(255, 255, 255);
                        break;
                    case 10:
                        actionUpdateColor.Invoke(255, 0, 0);
                        break;
                    case 11:
                        actionUpdateColor.Invoke(0, 255, 0);
                        break;
                    case 12:
                        actionUpdateColor.Invoke(0, 0, 255);
                        break;
                    case 13:
                        actionUpdateColor.Invoke(255, 255, 0);
                        break;
                    case 14:
                        actionUpdateColor.Invoke(255, 0, 255);
                        break;
                    case 15:
                        actionUpdateColor.Invoke(0, 255, 255);
                        break;
                    case 16:
                        actionUpdateColor.Invoke(128, 0, 0);
                        break;
                    case 17:
                        actionUpdateColor.Invoke(0, 128, 0);
                        break;
                    case 18:
                        actionUpdateColor.Invoke(0, 0, 128);
                        break;
                    case 19:
                        actionUpdateColor.Invoke(128, 128, 0);
                        break;
                    case 20:
                        actionUpdateColor.Invoke(128, 0, 128);
                        break;
                    case 21:
                        actionUpdateColor.Invoke(0, 128, 128);
                        break;
                    case 22:
                        actionUpdateColor.Invoke(192, 192, 192);
                        break;
                    case 23:
                        actionUpdateColor.Invoke(128, 128, 128);
                        break;
                    case 24:
                        actionUpdateColor.Invoke(153, 153, 255);
                        break;
                    case 25:
                        actionUpdateColor.Invoke(153, 51, 102);
                        break;
                    case 26:
                        actionUpdateColor.Invoke(255, 255, 204);
                        break;
                    case 27:
                        actionUpdateColor.Invoke(204, 255, 255);
                        break;
                    case 28:
                        actionUpdateColor.Invoke(102, 0, 102);
                        break;
                    case 29:
                        actionUpdateColor.Invoke(255, 128, 128);
                        break;
                    case 30:
                        actionUpdateColor.Invoke(0, 102, 204);
                        break;
                    case 31:
                        actionUpdateColor.Invoke(204, 204, 255);
                        break;
                    case 32:
                        actionUpdateColor.Invoke(0, 0, 128);
                        break;
                    case 33:
                        actionUpdateColor.Invoke(255, 0, 255);
                        break;
                    case 34:
                        actionUpdateColor.Invoke(255, 255, 0);
                        break;
                    case 35:
                        actionUpdateColor.Invoke(0, 255, 255);
                        break;
                    case 36:
                        actionUpdateColor.Invoke(128, 0, 128);
                        break;
                    case 37:
                        actionUpdateColor.Invoke(128, 0, 0);
                        break;
                    case 38:
                        actionUpdateColor.Invoke(0, 128, 128);
                        break;
                    case 39:
                        actionUpdateColor.Invoke(0, 0, 255);
                        break;
                    case 40:
                        actionUpdateColor.Invoke(0, 204, 255);
                        break;
                    case 41:
                        actionUpdateColor.Invoke(204, 255, 255);
                        break;
                    case 42:
                        actionUpdateColor.Invoke(204, 255, 204);
                        break;
                    case 43:
                        actionUpdateColor.Invoke(255, 255, 153);
                        break;
                    case 44:
                        actionUpdateColor.Invoke(153, 204, 255);
                        break;
                    case 45:
                        actionUpdateColor.Invoke(255, 153, 204);
                        break;
                    case 46:
                        actionUpdateColor.Invoke(204, 153, 255);
                        break;
                    case 47:
                        actionUpdateColor.Invoke(255, 204, 153);
                        break;
                    case 48:
                        actionUpdateColor.Invoke(51, 102, 255);
                        break;
                    case 49:
                        actionUpdateColor.Invoke(51, 204, 204);
                        break;
                    case 50:
                        actionUpdateColor.Invoke(153, 204, 0);
                        break;
                    case 51:
                        actionUpdateColor.Invoke(255, 204, 0);
                        break;
                    case 52:
                        actionUpdateColor.Invoke(255, 153, 0);
                        break;
                    case 53:
                        actionUpdateColor.Invoke(255, 102, 0);
                        break;
                    case 54:
                        actionUpdateColor.Invoke(102, 102, 153);
                        break;
                    case 55:
                        actionUpdateColor.Invoke(150, 150, 150);
                        break;
                    case 56:
                        actionUpdateColor.Invoke(0, 51, 102);
                        break;
                    case 57:
                        actionUpdateColor.Invoke(51, 153, 102);
                        break;
                    case 58:
                        actionUpdateColor.Invoke(0, 51, 0);
                        break;
                    case 59:
                        actionUpdateColor.Invoke(51, 51, 0);
                        break;
                    case 60:
                        actionUpdateColor.Invoke(153, 51, 0);
                        break;
                    case 61:
                        actionUpdateColor.Invoke(153, 51, 102);
                        break;
                    case 62:
                        actionUpdateColor.Invoke(51, 51, 153);
                        break;
                    case 63:
                        actionUpdateColor.Invoke(51, 51, 51);
                        break;
                    case 64:
                        actionUpdateColor.Invoke(128, 128, 128);
                        break;
                    case 65:
                        actionUpdateColor.Invoke(255, 255, 255);
                        break;
                    default:
                        return string.Empty;
                }
            }
            else if (color.Theme != null && color.Theme.HasValue && themeColors != null && color.Theme.Value >= 0 && color.Theme.Value < themeColors.Length)
            {
                DocumentFormat.OpenXml.Drawing.Color2Type themeColor = themeColors[color.Theme.Value];
                if (themeColor.RgbColorModelHex != null && themeColor.RgbColorModelHex.Val != null && themeColor.RgbColorModelHex.Val.HasValue)
                {
                    HexToRgba(themeColor.RgbColorModelHex.Val.Value, out result[0], out result[1], out result[2], out result[3]);
                }
                else if (themeColor.RgbColorModelPercentage != null)
                {
                    result[0] = themeColor.RgbColorModelPercentage.RedPortion != null && themeColor.RgbColorModelPercentage.RedPortion.HasValue ? (int)(themeColor.RgbColorModelPercentage.RedPortion.Value / 100000.0 * 255) : 0;
                    result[1] = themeColor.RgbColorModelPercentage.GreenPortion != null && themeColor.RgbColorModelPercentage.GreenPortion.HasValue ? (int)(themeColor.RgbColorModelPercentage.GreenPortion.Value / 100000.0 * 255) : 0;
                    result[2] = themeColor.RgbColorModelPercentage.BluePortion != null && themeColor.RgbColorModelPercentage.BluePortion.HasValue ? (int)(themeColor.RgbColorModelPercentage.BluePortion.Value / 100000.0 * 255) : 0;
                }
                else if (themeColor.HslColor != null)
                {
                    double hue = themeColor.HslColor.HueValue != null && themeColor.HslColor.HueValue.HasValue ? themeColor.HslColor.HueValue.Value / 60000.0 : 0;
                    double luminosity = themeColor.HslColor.LumValue != null && themeColor.HslColor.LumValue.HasValue ? themeColor.HslColor.LumValue.Value / 100000.0 : 0;
                    double saturation = themeColor.HslColor.SatValue != null && themeColor.HslColor.SatValue.HasValue ? themeColor.HslColor.SatValue.Value / 100000.0 : 0;
                    HlsToRgb(hue, luminosity, saturation, out result[0], out result[1], out result[2]);
                }
                else if (themeColor.SystemColor != null)
                {
                    if (themeColor.SystemColor.Val != null && themeColor.SystemColor.Val.HasValue && ThemeSystemColors.ContainsKey(themeColor.SystemColor.Val.Value))
                    {
                        double[] colorSystem = ThemeSystemColors[themeColor.SystemColor.Val.Value];
                        actionUpdateColor.Invoke(colorSystem[0], colorSystem[1], colorSystem[2]);
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
                else if (themeColor.PresetColor != null && themeColor.PresetColor.Val != null && themeColor.PresetColor.Val.HasValue && ThemePresetColors.ContainsKey(themeColor.PresetColor.Val.Value))
                {
                    double[] colorPreset = ThemePresetColors[themeColor.PresetColor.Val.Value];
                    actionUpdateColor.Invoke(colorPreset[0], colorPreset[1], colorPreset[2]);
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

            return $"{(result[3] >= 1 ? "rgb" : "rgba")}({Math.Round(result[0])}, {Math.Round(result[1])}, {Math.Round(result[2])}{(result[3] < 1 ? $", {Math.Round(result[3])}" : string.Empty)})";
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

            List<(string, DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties)> drawings = new List<(string, DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties)>();
            foreach (OpenXmlElement drawing in anchor.Descendants())
            {
                if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture && config.ConvertPictures && picture.BlipFill != null && picture.BlipFill.Blip != null && picture.BlipFill.Blip.Embed != null && picture.BlipFill.Blip.Embed.HasValue && worksheet.DrawingsPart != null && worksheet.DrawingsPart.GetPartById(picture.BlipFill.Blip.Embed.Value) is ImagePart imagePart)
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
                    string description = picture.NonVisualPictureProperties != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.HasValue ? $" alt=\"{picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.Value}\"" : string.Empty;
                    drawings.Add(($"<img src=\"data:{imagePart.ContentType};base64,{base64}\"{description}{{0}}/>", picture.NonVisualPictureProperties.NonVisualDrawingProperties, picture.ShapeProperties));
                }
                else if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape)
                {
                    //TODO: shapes
                    string text = shape.TextBody != null ? shape.TextBody.InnerText : string.Empty;
                    drawings.Add(($"<p{{0}}>{text}</p>", shape.NonVisualShapeProperties.NonVisualDrawingProperties, shape.ShapeProperties));
                }
            }
            foreach ((string, DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties) drawingInfo in drawings)
            {
                string leftActual = left;
                string topActual = top;
                string widthActual = width;
                string heightActual = height;
                string htmlTransforms = string.Empty;
                if (drawingInfo.Item3 != null && drawingInfo.Item3.Transform2D != null)
                {
                    if (drawingInfo.Item3.Transform2D.Offset != null)
                    {
                        if (left == "0" && drawingInfo.Item3.Transform2D.Offset.X != null && drawingInfo.Item3.Transform2D.Offset.X.HasValue)
                        {
                            htmlTransforms += $" translateX({RoundNumber(drawingInfo.Item3.Transform2D.Offset.X.Value / 914400.0 * 96, config.RoundingDigits)}px)";
                        }
                        if (top == "0" && drawingInfo.Item3.Transform2D.Offset.Y != null && drawingInfo.Item3.Transform2D.Offset.Y.HasValue)
                        {
                            htmlTransforms += $" translateY({RoundNumber(drawingInfo.Item3.Transform2D.Offset.Y.Value / 914400.0 * 96, config.RoundingDigits)}px)";
                        }
                    }
                    if (drawingInfo.Item3.Transform2D.Extents != null)
                    {
                        if (widthActual == "auto" && drawingInfo.Item3.Transform2D.Extents.Cx != null && drawingInfo.Item3.Transform2D.Extents.Cx.HasValue)
                        {
                            widthActual = $"{RoundNumber(drawingInfo.Item3.Transform2D.Extents.Cx.Value / 914400.0 * 96, config.RoundingDigits)}px";
                        }
                        if (heightActual == "auto" && drawingInfo.Item3.Transform2D.Extents.Cy != null && drawingInfo.Item3.Transform2D.Extents.Cy.HasValue)
                        {
                            heightActual = $"{RoundNumber(drawingInfo.Item3.Transform2D.Extents.Cy.Value / 914400.0 * 96, config.RoundingDigits)}px";
                        }
                    }
                    if (drawingInfo.Item3.Transform2D.Rotation != null && drawingInfo.Item3.Transform2D.Rotation.HasValue)
                    {
                        htmlTransforms += $" rotate(-{RoundNumber(drawingInfo.Item3.Transform2D.Rotation.Value, config.RoundingDigits)}deg)";
                    }
                    if (drawingInfo.Item3.Transform2D.HorizontalFlip != null && drawingInfo.Item3.Transform2D.HorizontalFlip.HasValue && drawingInfo.Item3.Transform2D.HorizontalFlip.Value)
                    {
                        htmlTransforms += $" scaleX(-1)";
                    }
                    if (drawingInfo.Item3.Transform2D.VerticalFlip != null && drawingInfo.Item3.Transform2D.VerticalFlip.HasValue && drawingInfo.Item3.Transform2D.VerticalFlip.Value)
                    {
                        htmlTransforms += $" scaleY(-1)";
                    }
                }

                bool isHidden = drawingInfo.Item2 != null && drawingInfo.Item2.Hidden != null && drawingInfo.Item2.Hidden.HasValue && drawingInfo.Item2.Hidden.Value;
                writer.Write($"\n{new string(' ', 8)}{drawingInfo.Item1.Replace("{0}", $" style=\"position: absolute; left: {leftActual}; top: {topActual}; width: {widthActual}; height: {heightActual};{(!string.IsNullOrEmpty(htmlTransforms) ? $" transform:{htmlTransforms};" : string.Empty)}{(isHidden ? " visibility: hidden;" : string.Empty)}\"")}");
            }
        }

        #endregion

        #region Private Fields

        private const string stringFormatNumber = "0.##############################";

        private static readonly Regex regexNumbers = new Regex(@"\d+", RegexOptions.Compiled);
        private static readonly Regex regexLetters = new Regex("[A-Za-z]+", RegexOptions.Compiled);

        private static readonly Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, double[]> ThemeSystemColors = new Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, double[]>()
        {
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveBorder, new double [3] { 180, 180, 180 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveCaption, new double [3] { 153, 180, 209 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ApplicationWorkspace, new double [3] { 171, 171, 171 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Background, new double [3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonFace, new double [3] { 240, 240, 240 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonHighlight, new double [3] { 0, 120, 215 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonShadow, new double [3] { 160, 160, 160 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonText, new double [3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.CaptionText, new double [3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientActiveCaption, new double [3] { 185, 209, 234 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientInactiveCaption, new double [3] { 215, 228, 242 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.GrayText, new double [3] { 109, 109, 109 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Highlight, new double [3] { 0, 120, 215 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.HighlightText, new double [3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.HotLight, new double [3] { 255, 165, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveBorder, new double [3] { 244, 247, 252 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaption, new double [3] { 191, 205, 219 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaptionText, new double [3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoBack, new double [3] { 255, 255, 225 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoText, new double [3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Menu, new double [3] { 240, 240, 240 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuBar, new double [3] { 240, 240, 240 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuHighlight, new double [3] { 0, 120, 215 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuText, new double [3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ScrollBar, new double [3] { 200, 200, 200 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDDarkShadow, new double [3] { 160, 160, 160 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDLight, new double [3] { 227, 227, 227 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Window, new double [3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowFrame, new double [3] { 100, 100, 100 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText, new double [3] { 0, 0, 0 } }
        };

        private static readonly Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, double[]> ThemePresetColors = new Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, double[]>()
        {
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.AliceBlue, new double[3] { 240, 248, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.AntiqueWhite, new double[3] { 250, 235, 215 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Aqua, new double[3] { 0, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Aquamarine, new double[3] { 127, 255, 212 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Azure, new double[3] { 240, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Beige, new double[3] { 245, 245, 220 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Bisque, new double[3] { 255, 228, 196 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Black, new double[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.BlanchedAlmond, new double[3] { 255, 235, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Blue, new double[3] { 0, 0, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.BlueViolet, new double[3] { 138, 43, 226 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Brown, new double[3] { 165, 42, 42 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.BurlyWood, new double[3] { 222, 184, 135 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.CadetBlue, new double[3] { 95, 158, 160 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Chartreuse, new double[3] { 127, 255, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Chocolate, new double[3] { 210, 105, 30 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Coral, new double[3] { 255, 127, 80 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.CornflowerBlue, new double[3] { 100, 149, 237 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Cornsilk, new double[3] { 255, 248, 220 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Crimson, new double[3] { 220, 20, 60 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Cyan, new double[3] { 0, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue, new double[3] { 0, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan, new double[3] { 0, 139, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod, new double[3] { 184, 134, 11 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray, new double[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen, new double[3] { 0, 100, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki, new double[3] { 189, 183, 107 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta, new double[3] { 139, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen, new double[3] { 85, 107, 47 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange, new double[3] { 255, 140, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid, new double[3] { 153, 50, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed, new double[3] { 139, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon, new double[3] { 233, 150, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen, new double[3] { 143, 188, 143 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue, new double[3] { 72, 61, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray, new double[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise, new double[3] { 0, 206, 209 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet, new double[3] { 148, 0, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepPink, new double[3] { 255, 20, 147 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepSkyBlue, new double[3] { 0, 191, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGray, new double[3] { 105, 105, 105 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DodgerBlue, new double[3] { 30, 144, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Firebrick, new double[3] { 178, 34, 34 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.FloralWhite, new double[3] { 255, 250, 240 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.ForestGreen, new double[3] { 34, 139, 34 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Fuchsia, new double[3] { 255, 0, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Gainsboro, new double[3] { 220, 220, 220 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.GhostWhite, new double[3] { 248, 248, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Gold, new double[3] { 255, 215, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Goldenrod, new double[3] { 218, 165, 32 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Gray, new double[3] { 128, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Green, new double[3] { 0, 128, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.GreenYellow, new double[3] { 173, 255, 47 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Honeydew, new double[3] { 240, 255, 240 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.HotPink, new double[3] { 255, 105, 180 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.IndianRed, new double[3] { 205, 92, 92 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Indigo, new double[3] { 75, 0, 130 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Ivory, new double[3] { 255, 255, 240 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Khaki, new double[3] { 240, 230, 140 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Lavender, new double[3] { 230, 230, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LavenderBlush, new double[3] { 255, 240, 245 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LawnGreen, new double[3] { 124, 252, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LemonChiffon, new double[3] { 255, 250, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue, new double[3] { 173, 216, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral, new double[3] { 240, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan, new double[3] { 224, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow, new double[3] { 250, 250, 210 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray, new double[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen, new double[3] { 144, 238, 144 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink, new double[3] { 255, 182, 193 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon, new double[3] { 255, 160, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen, new double[3] { 32, 178, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue, new double[3] { 135, 206, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray, new double[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue, new double[3] { 176, 196, 222 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow, new double[3] { 255, 255, 224 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Lime, new double[3] { 0, 255, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LimeGreen, new double[3] { 50, 205, 50 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Linen, new double[3] { 250, 240, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Magenta, new double[3] { 255, 0, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Maroon, new double[3] { 128, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MedAquamarine, new double[3] { 102, 205, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue, new double[3] { 0, 0, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid, new double[3] { 186, 85, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple, new double[3] { 147, 112, 219 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen, new double[3] { 60, 179, 113 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue, new double[3] { 123, 104, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen, new double[3] { 0, 250, 154 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise, new double[3] { 72, 209, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed, new double[3] { 199, 21, 133 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MidnightBlue, new double[3] { 25, 25, 112 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MintCream, new double[3] { 245, 255, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MistyRose, new double[3] { 255, 228, 225 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Moccasin, new double[3] { 255, 228, 181 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.NavajoWhite, new double[3] { 255, 222, 173 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Navy, new double[3] { 0, 0, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.OldLace, new double[3] { 253, 245, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Olive, new double[3] { 128, 128, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.OliveDrab, new double[3] { 107, 142, 35 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Orange, new double[3] { 255, 165, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.OrangeRed, new double[3] { 255, 69, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Orchid, new double[3] { 218, 112, 214 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGoldenrod, new double[3] { 238, 232, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGreen, new double[3] { 152, 251, 152 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleTurquoise, new double[3] { 175, 238, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleVioletRed, new double[3] { 219, 112, 147 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PapayaWhip, new double[3] { 255, 239, 213 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PeachPuff, new double[3] { 255, 218, 185 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Peru, new double[3] { 205, 133, 63 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Pink, new double[3] { 255, 192, 203 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Plum, new double[3] { 221, 160, 221 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PowderBlue, new double[3] { 176, 224, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Purple, new double[3] { 128, 0, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Red, new double[3] { 255, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.RosyBrown, new double[3] { 188, 143, 143 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.RoyalBlue, new double[3] { 65, 105, 225 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SaddleBrown, new double[3] { 139, 69, 19 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Salmon, new double[3] { 250, 128, 114 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SandyBrown, new double[3] { 244, 164, 96 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaGreen, new double[3] { 46, 139, 87 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaShell, new double[3] { 255, 245, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Sienna, new double[3] { 160, 82, 45 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Silver, new double[3] { 192, 192, 192 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SkyBlue, new double[3] { 135, 206, 235 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateBlue, new double[3] { 106, 90, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGray, new double[3] { 112, 128, 144 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Snow, new double[3] { 255, 250, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SpringGreen, new double[3] { 0, 255, 127 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SteelBlue, new double[3] { 70, 130, 180 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Tan, new double[3] { 210, 180, 140 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Teal, new double[3] { 0, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Thistle, new double[3] { 216, 191, 216 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Tomato, new double[3] { 255, 99, 71 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Turquoise, new double[3] { 64, 224, 208 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Violet, new double[3] { 238, 130, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Wheat, new double[3] { 245, 222, 179 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.White, new double[3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.WhiteSmoke, new double[3] { 245, 245, 245 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Yellow, new double[3] { 255, 255, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.YellowGreen, new double[3] { 154, 205, 50 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue2010, new double[3] { 0, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan2010, new double[3] { 0, 139, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod2010, new double[3] { 184, 134, 11 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray2010, new double[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey2010, new double[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen2010, new double[3] { 0, 100, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki2010, new double[3] { 189, 183, 107 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta2010, new double[3] { 139, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen2010, new double[3] { 85, 107, 47 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange2010, new double[3] { 255, 140, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid2010, new double[3] { 153, 50, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed2010, new double[3] { 139, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon2010, new double[3] { 233, 150, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen2010, new double[3] { 143, 188, 143 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue2010, new double[3] { 72, 61, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray2010, new double[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey2010, new double[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise2010, new double[3] { 0, 206, 209 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet2010, new double[3] { 148, 0, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue2010, new double[3] { 173, 216, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral2010, new double[3] { 240, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan2010, new double[3] { 224, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow2010, new double[3] { 250, 250, 210 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray2010, new double[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey2010, new double[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen2010, new double[3] { 144, 238, 144 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink2010, new double[3] { 255, 182, 193 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon2010, new double[3] { 255, 160, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen2010, new double[3] { 32, 178, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue2010, new double[3] { 135, 206, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray2010, new double[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey2010, new double[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue2010, new double[3] { 176, 196, 222 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow2010, new double[3] { 255, 255, 224 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumAquamarine2010, new double[3] { 102, 205, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue2010, new double[3] { 0, 0, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid2010, new double[3] { 186, 85, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple2010, new double[3] { 147, 112, 219 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen2010, new double[3] { 60, 179, 113 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue2010, new double[3] { 123, 104, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen2010, new double[3] { 0, 250, 154 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise2010, new double[3] { 72, 209, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed2010, new double[3] { 199, 21, 133 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey, new double[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGrey, new double[3] { 105, 105, 105 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey, new double[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Grey, new double[3] { 128, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey, new double[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey, new double[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGrey, new double[3] { 112, 128, 144 } }
        };

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
            this.PageTitle = "Conversion Result";
            this.PresetStyles = DefaultPresetStyles;
            this.ErrorMessage = DefaultErrorMessage;
            this.Encoding = System.Text.Encoding.UTF8;
            this.BufferSize = 65536;
            this.ConvertStyles = true;
            this.ConvertSizes = true;
            this.ConvertNumberFormats = true;
            this.ConvertPictures = true;
            this.ConvertSheetTitles = true;
            this.ConvertHiddenSheets = false;
            this.ConvertFirstSheetOnly = false;
            this.ConvertHtmlBodyOnly = false;
            this.UseHtmlStyleClasses = true;
            this.RoundingDigits = 2;
        }

        #region Public Fields

        /// <summary>
        /// Gets a new instance of <see cref="ConverterConfig"/> with default settings.
        /// </summary>
        public static ConverterConfig DefaultSettings { get { return new ConverterConfig(); } }

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
        /// Gets or sets whether to convert Xlsx texts with number formats to Html formatted texts.
        /// </summary>
        public bool ConvertNumberFormats { get; set; }

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
        /// Gets or sets whether to use the Html class attributes.
        /// </summary>
        public bool UseHtmlStyleClasses { get; set; }

        /// <summary>
        /// Gets or sets the number of digits to round the numbers to, or to not use rounding if the value is negative.
        /// </summary>
        public int RoundingDigits { get; set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Creates a cloned instance of the current <see cref="ConverterConfig"/> class.
        /// </summary>
        /// <returns>The cloned <see cref="ConverterConfig"/>.</returns>
        public ConverterConfig Clone()
        {
            return new ConverterConfig()
            {
                PageTitle = this.PageTitle,
                PresetStyles = this.PresetStyles,
                ErrorMessage = this.ErrorMessage,
                Encoding = this.Encoding,
                BufferSize = this.BufferSize,
                ConvertStyles = this.ConvertStyles,
                ConvertSizes = this.ConvertSizes,
                ConvertNumberFormats = this.ConvertNumberFormats,
                ConvertPictures = this.ConvertPictures,
                ConvertSheetTitles = this.ConvertSheetTitles,
                ConvertHiddenSheets = this.ConvertHiddenSheets,
                ConvertFirstSheetOnly = this.ConvertFirstSheetOnly,
                ConvertHtmlBodyOnly = this.ConvertHtmlBodyOnly,
                UseHtmlStyleClasses = this.UseHtmlStyleClasses,
                RoundingDigits = this.RoundingDigits
            };
        }

        #endregion
    }

    /// <summary>
    /// The progress callback event arguments of the Xlsx to Html converter.
    /// </summary>
    public class ConverterProgressCallbackEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConverterProgressCallbackEventArgs"/> class with specific progress.
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
        /// Gets the current progress in percentage that ranges from 0 to 100.
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
