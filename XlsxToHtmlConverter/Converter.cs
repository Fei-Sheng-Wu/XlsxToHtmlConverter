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
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once instead of using <see cref="FileStream"/>.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, bool loadIntoMemory = false)
        {
            ConvertXlsx(fileName, outputHtml, ConverterConfig.DefaultSettings, null, loadIntoMemory);
        }

        /// <summary>
        /// Converts a local Xlsx file to Html string with specific configuartions.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once instead of using <see cref="FileStream"/>.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, ConverterConfig config, bool loadIntoMemory = false)
        {
            ConvertXlsx(fileName, outputHtml, config, null, loadIntoMemory);
        }

        /// <summary>
        /// Converts a local Xlsx file to Html string with progress callback event.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once instead of using <see cref="FileStream"/>.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, EventHandler<ConverterProgressCallbackEventArgs> progressCallback, bool loadIntoMemory = false)
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
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once instead of using <see cref="FileStream"/>.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, ConverterConfig config, EventHandler<ConverterProgressCallbackEventArgs> progressCallback, bool loadIntoMemory = false)
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
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(inputXlsx, false))
            {
                ConvertXlsx(document, outputHtml, config, progressCallback);
            }
        }

        /// <summary>
        /// Converts a <see cref="SpreadsheetDocument"/> Xlsx file to Html string with specific configurations and progress callback event.
        /// </summary>
        /// <param name="inputDocument">The input <see cref="SpreadsheetDocument"/> object of the Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="config">The conversion configurations.</param>
        /// <param name="progressCallback">The progress callback event.</param>
        public static void ConvertXlsx(SpreadsheetDocument inputDocument, Stream outputHtml, ConverterConfig config, EventHandler<ConverterProgressCallbackEventArgs> progressCallback)
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
        {configClone.PresetStyles.Trim().Replace("\n", $"\n{new string(' ', 8)}")}
    </style>
</head>
<body>" : $"<style>\n{configClone.PresetStyles}\n</style>");

                WorkbookPart workbook = inputDocument.WorkbookPart;

                DocumentFormat.OpenXml.Drawing.Color2Type[] themes = null;
                if (workbook.ThemePart != null && workbook.ThemePart.Theme != null && workbook.ThemePart.Theme.ThemeElements != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme != null)
                {
                    themes = new DocumentFormat.OpenXml.Drawing.Color2Type[12] {
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
                (Dictionary<string, string>, string, uint)[] stylesheetCellFormats = new (Dictionary<string, string>, string, uint)[stylesheet != null && stylesheet.CellFormats != null ? stylesheet.CellFormats.ChildElements.Count : 0];
                for (int stylesheetFormatIndex = 0; stylesheetFormatIndex < stylesheetCellFormats.Length; stylesheetFormatIndex++)
                {
                    if (stylesheet.CellFormats.ChildElements[stylesheetFormatIndex] is CellFormat cellFormat)
                    {
                        Fill fill = (cellFormat.ApplyFill == null || !cellFormat.ApplyFill.HasValue || cellFormat.ApplyFill.Value) && cellFormat.FillId != null && cellFormat.FillId.HasValue && stylesheet.Fills != null && cellFormat.FillId.Value < stylesheet.Fills.ChildElements.Count ? (Fill)stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value] : null;
                        Font font = (cellFormat.ApplyFont == null || !cellFormat.ApplyFont.HasValue || cellFormat.ApplyFont.Value) && cellFormat.FontId != null && cellFormat.FontId.HasValue && stylesheet.Fonts != null && cellFormat.FontId.Value < stylesheet.Fonts.ChildElements.Count ? (Font)stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value] : null;
                        Border border = (cellFormat.ApplyBorder == null || !cellFormat.ApplyBorder.HasValue || cellFormat.ApplyBorder.Value) && cellFormat.BorderId != null && cellFormat.BorderId.HasValue && stylesheet.Borders != null && cellFormat.BorderId.Value < stylesheet.Borders.ChildElements.Count ? (Border)stylesheet.Borders.ChildElements[(int)cellFormat.BorderId.Value] : null;
                        string valueContainer = "{0}";
                        stylesheetCellFormats[stylesheetFormatIndex] = (GetCellFormat(fill, font, border, cellFormat.ApplyAlignment == null || !cellFormat.ApplyAlignment.HasValue || cellFormat.ApplyAlignment.Value ? cellFormat.Alignment : null, ref valueContainer, themes, configClone), valueContainer, cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.HasValue && (cellFormat.ApplyNumberFormat == null || !cellFormat.ApplyNumberFormat.HasValue || cellFormat.ApplyNumberFormat.Value) ? cellFormat.NumberFormatId.Value : 0);
                    }
                }
                (Dictionary<string, string>, string)[] stylesheetDifferentialFormats = new (Dictionary<string, string>, string)[stylesheet != null && stylesheet.DifferentialFormats != null ? stylesheet.DifferentialFormats.ChildElements.Count : 0];
                for (int stylesheetDifferentialFormatIndex = 0; stylesheetDifferentialFormatIndex < stylesheetDifferentialFormats.Length; stylesheetDifferentialFormatIndex++)
                {
                    if (stylesheet.DifferentialFormats.ChildElements[stylesheetDifferentialFormatIndex] is DifferentialFormat differentialFormat)
                    {
                        string valueContainer = "{0}";
                        stylesheetDifferentialFormats[stylesheetDifferentialFormatIndex] = (GetCellFormat(differentialFormat.Fill, differentialFormat.Font, differentialFormat.Border, differentialFormat.Alignment, ref valueContainer, themes, configClone), valueContainer);
                    }
                }
                Dictionary<uint, string[]> stylesheetNumberingFormats = new Dictionary<uint, string[]>();
                Dictionary<uint, string> stylesheetNumberingFormatsDateTimes = new Dictionary<uint, string>();
                Dictionary<string, (int, int, int, bool, int, int, int, int, bool, bool, string[])> stylesheetNumberingFormatsNumbers = new Dictionary<string, (int, int, int, bool, int, int, int, int, bool, bool, string[])>();
                if (configClone.ConvertNumberFormats && stylesheet != null && stylesheet.NumberingFormats != null)
                {
                    foreach (NumberingFormat numberingFormat in stylesheet.NumberingFormats.Elements<NumberingFormat>())
                    {
                        if (numberingFormat.NumberFormatId == null || !numberingFormat.NumberFormatId.HasValue)
                        {
                            continue;
                        }

                        string formatCode = numberingFormat.FormatCode != null && numberingFormat.FormatCode.HasValue ? System.Web.HttpUtility.HtmlDecode(numberingFormat.FormatCode.Value) : string.Empty;
                        List<string> formatCodeSplitted = new List<string>();
                        string formatCodeCurrent = string.Empty;
                        for (int i = 0; i < formatCode.Length; i++)
                        {
                            if (formatCode[i] == ';' && (i - 1 < 0 || formatCode[i - 1] != '\\'))
                            {
                                formatCodeSplitted.Add(formatCodeCurrent);
                                formatCodeCurrent = string.Empty;
                            }
                            else
                            {
                                formatCodeCurrent += formatCode[i];
                            }
                        }
                        formatCodeSplitted.Add(formatCodeCurrent);
                        stylesheetNumberingFormats[numberingFormat.NumberFormatId.Value] = formatCodeSplitted.ToArray();
                    }
                }

                SharedStringTable sharedStringTable = workbook.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() is SharedStringTablePart sharedStringTablePart ? sharedStringTablePart.SharedStringTable : null;
                (string, string)[] sharedStrings = new (string, string)[sharedStringTable != null ? sharedStringTable.ChildElements.Count : 0];
                for (int sharedStringIndex = 0; sharedStringIndex < sharedStrings.Length; sharedStringIndex++)
                {
                    if (sharedStringTable.ChildElements[sharedStringIndex] is SharedStringItem sharedString)
                    {
                        string sharedStringValue = string.Empty;
                        string sharedStringValueRaw = string.Empty;
                        if (configClone.ConvertStyles)
                        {
                            foreach (OpenXmlElement sharedStringElement in sharedString.Elements())
                            {
                                if (sharedStringElement is Text text)
                                {
                                    sharedStringValue += GetEscapedString(text.Text);
                                    sharedStringValueRaw += text.Text;
                                }
                                else if (sharedStringElement is Run run && run.Text != null)
                                {
                                    string runValueContainer = "{0}";
                                    Dictionary<string, string> runStyles = null;
                                    if (run.RunProperties is RunProperties runProperties)
                                    {
                                        GetFont(runProperties, ref runStyles, ref runValueContainer, themes, configClone);
                                    }
                                    sharedStringValue += $"<span{(runStyles != null && runStyles.Count > 0 ? $" style=\"{GetAttributesString(runStyles, false, -1)}\"" : string.Empty)}>{runValueContainer.Replace("{0}", GetEscapedString(run.Text.Text))}</span>";
                                    sharedStringValueRaw += run.Text.Text;
                                }
                            }
                        }
                        else
                        {
                            string text = sharedString.Text != null ? sharedString.Text.Text : string.Empty;
                            sharedStringValue = GetEscapedString(text);
                            sharedStringValueRaw = text;
                        }
                        sharedStrings[sharedStringIndex] = (sharedStringValue, sharedStringValueRaw != sharedStringValue ? sharedStringValueRaw : string.Empty);
                    }
                }

                IEnumerable<Sheet> sheets = workbook.Workbook != null && workbook.Workbook.Sheets != null ? workbook.Workbook.Sheets.Elements<Sheet>() : Enumerable.Empty<Sheet>();
                int sheetIndex = 0;
                int sheetsCount = configClone.ConvertFirstSheetOnly ? Math.Min(1, sheets.Count()) : sheets.Count();
                foreach (Sheet sheet in sheets)
                {
                    sheetIndex++;
                    if ((configClone.ConvertFirstSheetOnly && sheetIndex > 1) || sheet.Id == null || !sheet.Id.HasValue || !(workbook.GetPartById(sheet.Id.Value) is WorksheetPart worksheetPart) || (!configClone.ConvertHiddenSheets && sheet.State != null && sheet.State.HasValue && sheet.State.Value != SheetStateValues.Visible))
                    {
                        continue;
                    }
                    string sheetName = sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : string.Empty;

                    foreach (TableDefinitionPart tableDefinitionPart in worksheetPart.TableDefinitionParts ?? Enumerable.Empty<TableDefinitionPart>())
                    {
                        //TODO: tables
                        if (tableDefinitionPart.Table == null || tableDefinitionPart.Table.Reference == null || !tableDefinitionPart.Table.Reference.HasValue)
                        {
                            continue;
                        }

                        GetReferenceRange(tableDefinitionPart.Table.Reference.Value, out int tableFromColumn, out int tableFromRow, out int tableToColumn, out int tableToRow);
                    }

                    Dictionary<int, double> drawingColumnMarkers = new Dictionary<int, double>();
                    Dictionary<int, double> drawingRowMarkers = new Dictionary<int, double>();
                    foreach (OpenXmlElement drawing in worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null ? worksheetPart.DrawingsPart.WorksheetDrawing.Elements() : Enumerable.Empty<OpenXmlElement>())
                    {
                        if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor && oneCellAnchor.FromMarker != null)
                        {
                            if (oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) && !drawingColumnMarkers.ContainsKey(columnId))
                            {
                                drawingColumnMarkers[columnId] = double.NaN;
                            }
                            if (oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) && !drawingRowMarkers.ContainsKey(rowId))
                            {
                                drawingRowMarkers[rowId] = double.NaN;
                            }
                        }
                        else if (drawing is DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor)
                        {
                            if (twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) && !drawingColumnMarkers.ContainsKey(fromColumnId))
                            {
                                drawingColumnMarkers[fromColumnId] = double.NaN;
                            }
                            if (twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) && !drawingRowMarkers.ContainsKey(fromRowId))
                            {
                                drawingRowMarkers[fromRowId] = double.NaN;
                            }
                            if (twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) && !drawingColumnMarkers.ContainsKey(toColumnId))
                            {
                                drawingColumnMarkers[toColumnId] = double.NaN;
                            }
                            if (twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) && !drawingRowMarkers.ContainsKey(toRowId))
                            {
                                drawingRowMarkers[toRowId] = double.NaN;
                            }
                        }
                    }

                    if (!(worksheetPart.Worksheet is Worksheet worksheet))
                    {
                        continue;
                    }
                    SheetData sheetData = null;
                    bool isSheetDimensionSet = false;
                    int sheetDimensionFromColumn = 0;
                    int sheetDimensionFromRow = 0;
                    int sheetDimensionToColumn = 0;
                    int sheetDimensionToRow = 0;
                    string sheetTabColor = string.Empty;
                    double columnWidthDefault = 8.43;
                    double rowHeightDefault = 20;
                    List<(int, int, double)> columns = new List<(int, int, double)>();
                    Dictionary<(int, int), object> mergeCells = new Dictionary<(int, int), object>();
                    List<(ConditionalFormatting, List<(int, int, int, int)>, IEnumerable<ConditionalFormattingRule>)> conditionalFormattings = new List<(ConditionalFormatting, List<(int, int, int, int)>, IEnumerable<ConditionalFormattingRule>)>();
                    foreach (OpenXmlElement worksheetElement in worksheet.Elements())
                    {
                        if (worksheetElement is SheetData worksheetData)
                        {
                            sheetData = worksheetData;
                        }
                        else if (worksheetElement is SheetDimension worksheetDimension && worksheetDimension.Reference != null && worksheetDimension.Reference.HasValue)
                        {
                            isSheetDimensionSet = true;
                            GetReferenceRange(worksheetDimension.Reference.Value, out sheetDimensionFromColumn, out sheetDimensionFromRow, out sheetDimensionToColumn, out sheetDimensionToRow);
                        }
                        else if (worksheetElement is SheetProperties worksheetProperties && configClone.ConvertSheetTitles && worksheetProperties.TabColor != null)
                        {
                            GetColor(worksheetProperties.TabColor, out sheetTabColor, themes, configClone);
                        }
                        else if (worksheetElement is SheetFormatProperties worksheetFormatProperties)
                        {
                            columnWidthDefault = worksheetFormatProperties.DefaultColumnWidth != null && worksheetFormatProperties.DefaultColumnWidth.HasValue ? worksheetFormatProperties.DefaultColumnWidth.Value : (worksheetFormatProperties.BaseColumnWidth != null && worksheetFormatProperties.BaseColumnWidth.HasValue ? worksheetFormatProperties.BaseColumnWidth.Value : columnWidthDefault);
                            rowHeightDefault = worksheetFormatProperties.DefaultRowHeight != null && worksheetFormatProperties.DefaultRowHeight.HasValue ? RoundNumber(worksheetFormatProperties.DefaultRowHeight.Value / 72 * 96, configClone.RoundingDigits) : rowHeightDefault;
                        }
                        else if (worksheetElement is Columns worksheetColumns && configClone.ConvertSizes)
                        {
                            foreach (Column column in worksheetColumns.Elements<Column>())
                            {
                                bool isHidden = (column.Collapsed != null && (!column.Collapsed.HasValue || column.Collapsed.Value)) || (column.Hidden != null && (!column.Hidden.HasValue || column.Hidden.Value));
                                if ((column.Width != null && column.Width.HasValue && (column.CustomWidth == null || !column.CustomWidth.HasValue || column.CustomWidth.Value)) || isHidden)
                                {
                                    columns.Add((column.Min != null && column.Min.HasValue ? (int)column.Min.Value : int.MinValue, column.Max != null && column.Max.HasValue ? (int)column.Max.Value : int.MaxValue, isHidden ? 0 : column.Width.Value));
                                }
                            }
                        }
                        else if (worksheetElement is MergeCells worksheetMergeCells)
                        {
                            foreach (MergeCell mergeCell in worksheetMergeCells.Elements<MergeCell>())
                            {
                                if (mergeCell.Reference == null || !mergeCell.Reference.HasValue)
                                {
                                    continue;
                                }

                                GetReferenceRange(mergeCell.Reference.Value, out int mergeCellFromColumn, out int mergeCellFromRow, out int mergeCellToColumn, out int mergeCellToRow);
                                mergeCells[(mergeCellFromColumn, mergeCellFromRow)] = (mergeCellToColumn - mergeCellFromColumn + 1, mergeCellToRow - mergeCellFromRow + 1);
                                for (int i = mergeCellFromColumn; i <= mergeCellToColumn; i++)
                                {
                                    for (int j = mergeCellFromRow; j <= mergeCellToRow; j++)
                                    {
                                        if (!mergeCells.ContainsKey((i, j)))
                                        {
                                            mergeCells[(i, j)] = null;
                                        }
                                    }
                                }
                            }
                        }
                        else if (worksheetElement is ConditionalFormatting worksheetConditionalFormatting)
                        {
                            if (worksheetConditionalFormatting.SequenceOfReferences != null && worksheetConditionalFormatting.SequenceOfReferences.HasValue)
                            {
                                List<(int, int, int, int)> sequence = new List<(int, int, int, int)>();
                                foreach (string references in worksheetConditionalFormatting.SequenceOfReferences.Items)
                                {
                                    GetReferenceRange(references, out int referenceFromColumn, out int referenceFromRow, out int referenceToColumn, out int referenceToRow);
                                    sequence.Add((referenceFromColumn, referenceFromRow, referenceToColumn, referenceToRow));
                                }
                                conditionalFormattings.Add((worksheetConditionalFormatting, sequence, worksheetConditionalFormatting.Elements<ConditionalFormattingRule>()));
                            }
                        }
                    }

                    writer.Write($"\n{new string(' ', 4)}<h5{(!string.IsNullOrEmpty(sheetTabColor) ? $" style=\"border-bottom-color: {sheetTabColor};\"" : string.Empty)}>{(!string.IsNullOrEmpty(sheetName) ? sheetName : "Untitled Sheet")}</h5>");
                    writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                    writer.Write($"\n{new string(' ', 8)}<table>");

                    if (!isSheetDimensionSet)
                    {
                        sheetDimensionFromColumn = 1;
                        sheetDimensionFromRow = 1;
                        sheetDimensionToColumn = 1;
                        sheetDimensionToRow = 1;
                        foreach (Cell cell in sheetData != null ? sheetData.Elements<Row>().SelectMany(x => x.Elements<Cell>()) : Enumerable.Empty<Cell>())
                        {
                            if (cell.CellReference != null && cell.CellReference.HasValue)
                            {
                                sheetDimensionToColumn = Math.Max(sheetDimensionToColumn, GetColumnIndex(cell.CellReference.Value));
                                sheetDimensionToRow = Math.Max(sheetDimensionToRow, GetRowIndex(cell.CellReference.Value));
                            }
                        }
                    }

                    double[] columnWidths = new double[sheetDimensionToColumn - sheetDimensionFromColumn + 1];
                    for (int columnWidthIndex = 0; columnWidthIndex < columnWidths.Length; columnWidthIndex++)
                    {
                        columnWidths[columnWidthIndex] = columnWidthDefault;
                    }
                    if (configClone.ConvertSizes)
                    {
                        foreach ((int, int, double) columnInfo in columns)
                        {
                            for (int i = Math.Max(sheetDimensionFromColumn, columnInfo.Item1); i <= Math.Min(sheetDimensionToColumn, columnInfo.Item2); i++)
                            {
                                columnWidths[i - sheetDimensionFromColumn] = columnInfo.Item3;
                            }
                        }
                        double columnWidthsTotal = columnWidths.Sum();
                        double columbWidthsAccumulation = 0;
                        for (int columnWidthIndex = 0; columnWidthIndex < columnWidths.Length; columnWidthIndex++)
                        {
                            columnWidths[columnWidthIndex] = RoundNumber(columnWidths[columnWidthIndex] / columnWidthsTotal * 100, configClone.RoundingDigits);
                            columbWidthsAccumulation += columnWidths[columnWidthIndex];
                            if (drawingColumnMarkers.ContainsKey(sheetDimensionFromColumn + columnWidthIndex))
                            {
                                drawingColumnMarkers[sheetDimensionFromColumn + columnWidthIndex] = RoundNumber(columbWidthsAccumulation, configClone.RoundingDigits);
                            }
                        }
                    }

                    int rowIndex = sheetDimensionFromRow;
                    double rowHeightsAccumulation = 0;
                    foreach (Row row in sheetData != null ? sheetData.Elements<Row>() : Enumerable.Empty<Row>())
                    {
                        rowIndex++;
                        if (row.RowIndex != null && row.RowIndex.HasValue)
                        {
                            if (row.RowIndex.Value < sheetDimensionFromRow || row.RowIndex.Value > sheetDimensionToRow)
                            {
                                continue;
                            }

                            for (int additionalRowIndex = rowIndex; additionalRowIndex < row.RowIndex.Value; additionalRowIndex++)
                            {
                                if (configClone.ConvertSizes)
                                {
                                    rowHeightsAccumulation += rowHeightDefault;
                                    if (drawingRowMarkers.ContainsKey(additionalRowIndex))
                                    {
                                        drawingRowMarkers[additionalRowIndex] = RoundNumber(rowHeightsAccumulation, configClone.RoundingDigits);
                                    }
                                }

                                writer.Write($"\n{new string(' ', 12)}<tr>");
                                for (int additionalColumnIndex = 0; additionalColumnIndex < columnWidths.Length; additionalColumnIndex++)
                                {
                                    writer.Write($"\n{new string(' ', 16)}<td style=\"width: {(configClone.ConvertSizes ? $"{GetInvariantNumber(columnWidths[additionalColumnIndex])}%" : "auto")}; height: {(configClone.ConvertSizes ? $"{GetInvariantNumber(rowHeightDefault)}px" : "auto")};\"></td>");
                                }
                                writer.Write($"\n{new string(' ', 12)}</tr>");
                            }
                            rowIndex = (int)row.RowIndex.Value;
                        }
                        double cellHeightActual = (row.CustomHeight == null || !row.CustomHeight.HasValue || row.CustomHeight.Value) && row.Height != null && row.Height.HasValue ? RoundNumber(row.Height.Value / 72 * 96, configClone.RoundingDigits) : rowHeightDefault;
                        if (configClone.ConvertSizes)
                        {
                            rowHeightsAccumulation += cellHeightActual;
                            if (drawingRowMarkers.ContainsKey(rowIndex))
                            {
                                drawingRowMarkers[rowIndex] = RoundNumber(rowHeightsAccumulation, configClone.RoundingDigits);
                            }
                        }

                        writer.Write($"\n{new string(' ', 12)}<tr>");

                        Cell[] cells = new Cell[columnWidths.Length];
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            if (cell.CellReference != null && cell.CellReference.HasValue)
                            {
                                int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
                                if (cellColumnIndex >= sheetDimensionFromColumn && cellColumnIndex <= sheetDimensionToColumn)
                                {
                                    cells[cellColumnIndex - sheetDimensionFromColumn] = cell;
                                }
                            }
                        }
                        for (int cellIndex = sheetDimensionFromColumn; cellIndex <= sheetDimensionToColumn; cellIndex++)
                        {
                            string cellColumnName = string.Empty;
                            int cellColumnIndex = cellIndex;
                            while (cellColumnIndex > 0)
                            {
                                int modulo = (cellColumnIndex - 1) % 26;
                                cellColumnName = (char)(65 + modulo) + cellColumnName;
                                cellColumnIndex = (cellColumnIndex - modulo) / 26;
                            }
                            cells[cellIndex - sheetDimensionFromColumn] = cells[cellIndex - sheetDimensionFromColumn] ?? new Cell() { CellValue = new CellValue(string.Empty) };
                            cells[cellIndex - sheetDimensionFromColumn].CellReference = $"{cellColumnName}{rowIndex}";
                        }

                        int columnIndex = sheetDimensionFromColumn;
                        foreach (Cell cell in cells)
                        {
                            columnIndex = GetColumnIndex(cell.CellReference.Value);
                            double cellWidthActual = columnWidths[columnIndex - sheetDimensionFromColumn];

                            int columnSpanned = 1;
                            int rowSpanned = 1;
                            if (mergeCells.ContainsKey((columnIndex, rowIndex)))
                            {
                                if (!(mergeCells[(columnIndex, rowIndex)] is ValueTuple<int, int> mergeCellInfo))
                                {
                                    continue;
                                }
                                columnSpanned = mergeCellInfo.Item1;
                                rowSpanned = mergeCellInfo.Item2;
                            }

                            int styleIndex = cell.StyleIndex != null && cell.StyleIndex.HasValue ? (int)cell.StyleIndex.Value : (row.StyleIndex != null && row.StyleIndex.HasValue ? (int)row.StyleIndex.Value : -1);
                            Dictionary<string, string> cellStyles = new Dictionary<string, string>();
                            string cellValueContainer = "{0}";
                            if (configClone.ConvertStyles && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length)
                            {
                                cellStyles = !configClone.UseHtmlStyleClasses && stylesheetCellFormats[styleIndex].Item1 != null ? GetJoinedAttributes(cellStyles, stylesheetCellFormats[styleIndex].Item1) : cellStyles;
                                cellValueContainer = !string.IsNullOrEmpty(stylesheetCellFormats[styleIndex].Item2) ? cellValueContainer.Replace("{0}", stylesheetCellFormats[styleIndex].Item2) : cellValueContainer;
                            }

                            string cellValue = string.Empty;
                            string cellValueRaw = string.Empty;
                            bool isCellValueNumber = false;
                            if (cell.CellValue != null)
                            {
                                bool isCellValueSharedString = false;
                                if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.SharedString && int.TryParse(cell.CellValue.Text, out int sharedStringId) && sharedStringId >= 0 && sharedStringId < sharedStrings.Length)
                                {
                                    isCellValueSharedString = true;
                                    cellValue = sharedStrings[sharedStringId].Item1;
                                    cellValueRaw = sharedStrings[sharedStringId].Item2;
                                    cellValueRaw = string.IsNullOrEmpty(cellValueRaw) ? cellValue : cellValueRaw;
                                }
                                else
                                {
                                    cellValue = cell.CellValue.Text;
                                    cellValueRaw = cellValue;
                                }
                                isCellValueNumber = double.TryParse(cellValueRaw, out double cellValueNumber);

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
                                                string[] numberFormatCodeParts = stylesheetNumberingFormats[stylesheetCellFormats[styleIndex].Item3];
                                                if (numberFormatCodeParts.Length > 1 && isCellValueNumber)
                                                {
                                                    int indexComponent = cellValueNumber > 0 || (numberFormatCodeParts.Length == 2 && cellValueNumber == 0) ? 0 : (cellValueNumber < 0 ? 1 : (numberFormatCodeParts.Length > 2 ? 2 : -1));
                                                    numberFormatCode = indexComponent >= 0 ? numberFormatCodeParts[indexComponent] : numberFormatCode;
                                                }
                                                else if (numberFormatCodeParts.Length > 0)
                                                {
                                                    numberFormatCode = numberFormatCodeParts[numberFormatCodeParts.Length - 1];
                                                }
                                            }
                                            break;
                                    }
                                }

                                if (!string.IsNullOrEmpty(numberFormatCode))
                                {
                                    if ((isNumberFormatDefaultDateTime || (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Date)) && isCellValueNumber)
                                    {
                                        if (!isNumberFormatDefaultDateTime && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length && stylesheetNumberingFormatsDateTimes.ContainsKey(stylesheetCellFormats[styleIndex].Item3))
                                        {
                                            numberFormatCode = stylesheetNumberingFormatsDateTimes[stylesheetCellFormats[styleIndex].Item3];
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
                                                stylesheetNumberingFormatsDateTimes[stylesheetCellFormats[styleIndex].Item3] = numberFormatCodeNew;
                                            }
                                            numberFormatCode = numberFormatCodeNew;
                                        }
                                        cellValue = GetEscapedString(DateTime.FromOADate(cellValueNumber).Date.ToString(numberFormatCode));
                                    }
                                    else
                                    {
                                        cellValue = GetEscapedString(GetFormattedNumber(cellValueRaw, numberFormatCode.Trim(), ref stylesheetNumberingFormatsNumbers, x =>
                                        {
                                            for (int i = configClone.ConvertStyles && isCellValueNumber ? 0 : x.Length; i < x.Length; i++)
                                            {
                                                string conditionColor = string.Empty;
                                                switch (x[i].ToLower())
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
                                                if (i + 1 < x.Length && !char.IsLetter(x[i + 1][0]))
                                                {
                                                    i++;
                                                    if (x[i].StartsWith("="))
                                                    {
                                                        isConditionMet = double.TryParse(x[i].Substring(1, x[i].Length - 1), out double conditionValue) && cellValueNumber == conditionValue;
                                                    }
                                                    else if (x[i].StartsWith("<>"))
                                                    {
                                                        isConditionMet = double.TryParse(x[i].Substring(2, x[i].Length - 2), out double conditionValue) && cellValueNumber != conditionValue;
                                                    }
                                                    else if (x[i].StartsWith(">="))
                                                    {
                                                        isConditionMet = double.TryParse(x[i].Substring(2, x[i].Length - 2), out double conditionValue) && cellValueNumber >= conditionValue;
                                                    }
                                                    else if (x[i].StartsWith("<="))
                                                    {
                                                        isConditionMet = double.TryParse(x[i].Substring(2, x[i].Length - 2), out double conditionValue) && cellValueNumber <= conditionValue;
                                                    }
                                                    else if (x[i].StartsWith(">"))
                                                    {
                                                        isConditionMet = double.TryParse(x[i].Substring(1, x[i].Length - 1), out double conditionValue) && cellValueNumber > conditionValue;
                                                    }
                                                    else if (x[i].StartsWith("<"))
                                                    {
                                                        isConditionMet = double.TryParse(x[i].Substring(1, x[i].Length - 1), out double conditionValue) && cellValueNumber < conditionValue;
                                                    }
                                                }
                                                if (!string.IsNullOrEmpty(conditionColor) && isConditionMet)
                                                {
                                                    cellStyles["color"] = conditionColor;
                                                }
                                            }
                                        }));
                                    }
                                }
                                else if (!isCellValueSharedString)
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
                                        cellStyles["text-align"] = "center";
                                    }
                                    else if (cell.DataType.Value == CellValues.Date || cell.DataType.Value == CellValues.Number)
                                    {
                                        cellStyles["text-align"] = "right";
                                    }
                                }
                                else if (isCellValueNumber)
                                {
                                    cellStyles["text-align"] = "right";
                                }

                                int differentialStyleIndex = -1;
                                foreach ((ConditionalFormatting, List<(int, int, int, int)>, IEnumerable<ConditionalFormattingRule>) conditionalFormatting in conditionalFormattings)
                                {
                                    if (!conditionalFormatting.Item2.Any(x => columnIndex >= x.Item1 && columnIndex <= x.Item3 && rowIndex >= x.Item2 && rowIndex <= x.Item4))
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
                                                isConditionMet = GetFormulaCondition(cellValueRaw, formattingRule.Elements<Formula>(), 1, x => x[0] > x[1]);
                                            }
                                            else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.GreaterThanOrEqual)
                                            {
                                                isConditionMet = GetFormulaCondition(cellValueRaw, formattingRule.Elements<Formula>(), 1, x => x[0] >= x[1]);
                                            }
                                            else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.LessThan)
                                            {
                                                isConditionMet = GetFormulaCondition(cellValueRaw, formattingRule.Elements<Formula>(), 1, x => x[0] < x[1]);
                                            }
                                            else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.LessThanOrEqual)
                                            {
                                                isConditionMet = GetFormulaCondition(cellValueRaw, formattingRule.Elements<Formula>(), 1, x => x[0] <= x[1]);
                                            }
                                            else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.Between)
                                            {
                                                isConditionMet = GetFormulaCondition(cellValueRaw, formattingRule.Elements<Formula>(), 2, x => x[0] >= Math.Min(x[1], x[2]) && x[0] <= Math.Max(x[1], x[2]));
                                            }
                                            else if (formattingRule.Operator.Value == ConditionalFormattingOperatorValues.NotBetween)
                                            {
                                                isConditionMet = GetFormulaCondition(cellValueRaw, formattingRule.Elements<Formula>(), 2, x => x[0] < Math.Min(x[1], x[2]) || x[0] > Math.Max(x[1], x[2]));
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
                                    cellStyles = GetJoinedAttributes(cellStyles, stylesheetDifferentialFormats[differentialStyleIndex].Item1);
                                    cellValueContainer = cellValueContainer.Replace("{0}", stylesheetDifferentialFormats[differentialStyleIndex].Item2);
                                }
                            }

                            writer.Write($"\n{new string(' ', 16)}<td{(columnSpanned > 1 ? $" colspan=\"{columnSpanned}\"" : string.Empty)}{(rowSpanned > 1 ? $" rowspan=\"{rowSpanned}\"" : string.Empty)}{(configClone.UseHtmlStyleClasses && styleIndex >= 0 && styleIndex < stylesheetCellFormats.Length ? $" class=\"format-{styleIndex}\"" : string.Empty)} style=\"width: {(configClone.ConvertSizes && columnSpanned <= 1 ? $"{GetInvariantNumber(cellWidthActual)}%" : "auto")}; height: {(configClone.ConvertSizes && rowSpanned <= 1 ? $"{GetInvariantNumber(cellHeightActual)}px" : "auto")};{GetAttributesString(cellStyles, true, -1)}\">{cellValueContainer.Replace("{0}", cellValue)}</td>");
                        }

                        writer.Write($"\n{new string(' ', 12)}</tr>");

                        progressCallback?.Invoke(inputDocument, new ConverterProgressCallbackEventArgs(sheetIndex, sheetsCount, rowIndex - sheetDimensionFromRow + 1, sheetDimensionToRow - sheetDimensionFromRow + 1));
                    }

                    writer.Write($"\n{new string(' ', 8)}</table>");

                    if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                    {
                        foreach (OpenXmlElement drawingElement in worksheetPart.DrawingsPart.WorksheetDrawing.Elements())
                        {
                            string drawing = string.Empty;
                            if (drawingElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor anchorAbsolute)
                            {
                                string left = anchorAbsolute.Position != null && anchorAbsolute.Position.X != null && anchorAbsolute.Position.X.HasValue ? $"{GetInvariantNumber(RoundNumber(anchorAbsolute.Position.X.Value / 914400.0 * 96, configClone.RoundingDigits))}px" : "0";
                                string top = anchorAbsolute.Position != null && anchorAbsolute.Position.Y != null && anchorAbsolute.Position.Y.HasValue ? $"{GetInvariantNumber(RoundNumber(anchorAbsolute.Position.Y.Value / 914400.0 * 96, configClone.RoundingDigits))}px" : "0";
                                string width = anchorAbsolute.Extent != null && anchorAbsolute.Extent.Cx != null && anchorAbsolute.Extent.Cx.HasValue ? $"{GetInvariantNumber(RoundNumber(anchorAbsolute.Extent.Cx.Value / 914400.0 * 96, configClone.RoundingDigits))}px" : "auto";
                                string height = anchorAbsolute.Extent != null && anchorAbsolute.Extent.Cy != null && anchorAbsolute.Extent.Cy.HasValue ? $"{GetInvariantNumber(RoundNumber(anchorAbsolute.Extent.Cy.Value / 914400.0 * 96, configClone.RoundingDigits))}px" : "auto";
                                drawing = GetDrawing(anchorAbsolute, left, top, width, height, worksheetPart.DrawingsPart, themes, configClone);
                            }
                            else if (drawingElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor anchorOneCell)
                            {
                                double left = configClone.ConvertSizes && anchorOneCell.FromMarker != null && anchorOneCell.FromMarker.ColumnId != null && int.TryParse(anchorOneCell.FromMarker.ColumnId.Text, out int columnId) && drawingColumnMarkers.ContainsKey(columnId) ? drawingColumnMarkers[columnId] : double.NaN;
                                double leftOffset = anchorOneCell.FromMarker.ColumnOffset != null && int.TryParse(anchorOneCell.FromMarker.ColumnOffset.Text, out int columnOffset) ? RoundNumber(columnOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                double top = configClone.ConvertSizes && anchorOneCell.FromMarker != null && anchorOneCell.FromMarker.RowId != null && int.TryParse(anchorOneCell.FromMarker.RowId.Text, out int rowId) && drawingRowMarkers.ContainsKey(rowId) ? drawingRowMarkers[rowId] : double.NaN;
                                double topOffset = anchorOneCell.FromMarker.RowOffset != null && int.TryParse(anchorOneCell.FromMarker.RowOffset.Text, out int rowOffset) ? RoundNumber(rowOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                string width = anchorOneCell.Extent != null && anchorOneCell.Extent.Cx != null && anchorOneCell.Extent.Cx.HasValue ? $"{GetInvariantNumber(RoundNumber(anchorOneCell.Extent.Cx.Value / 914400.0 * 96, configClone.RoundingDigits))}px" : "auto";
                                string height = anchorOneCell.Extent != null && anchorOneCell.Extent.Cy != null && anchorOneCell.Extent.Cy.HasValue ? $"{GetInvariantNumber(RoundNumber(anchorOneCell.Extent.Cy.Value / 914400.0 * 96, configClone.RoundingDigits))}px" : "auto";
                                drawing = GetDrawing(anchorOneCell, !double.IsNaN(left) ? $"calc({GetInvariantNumber(left)}% + {GetInvariantNumber(leftOffset)}px)" : "0", !double.IsNaN(top) ? $"{GetInvariantNumber(RoundNumber(top + topOffset, configClone.RoundingDigits))}px" : "0", width, height, worksheetPart.DrawingsPart, themes, configClone);
                            }
                            else if (drawingElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor anchorTwoCell)
                            {
                                double left = configClone.ConvertSizes && anchorTwoCell.FromMarker != null && anchorTwoCell.FromMarker.ColumnId != null && int.TryParse(anchorTwoCell.FromMarker.ColumnId.Text, out int fromColumnId) && drawingColumnMarkers.ContainsKey(fromColumnId) ? drawingColumnMarkers[fromColumnId] : double.NaN;
                                double leftOffset = anchorTwoCell.FromMarker.ColumnOffset != null && int.TryParse(anchorTwoCell.FromMarker.ColumnOffset.Text, out int fromMarkerColumnOffset) ? RoundNumber(fromMarkerColumnOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                double top = configClone.ConvertSizes && anchorTwoCell.FromMarker != null && anchorTwoCell.FromMarker.RowId != null && int.TryParse(anchorTwoCell.FromMarker.RowId.Text, out int fromRowId) && drawingRowMarkers.ContainsKey(fromRowId) ? drawingRowMarkers[fromRowId] : double.NaN;
                                double topOffset = anchorTwoCell.FromMarker.RowOffset != null && int.TryParse(anchorTwoCell.FromMarker.RowOffset.Text, out int fromMarkerRowOffset) ? RoundNumber(fromMarkerRowOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                double right = configClone.ConvertSizes && anchorTwoCell.ToMarker != null && anchorTwoCell.ToMarker.ColumnId != null && int.TryParse(anchorTwoCell.ToMarker.ColumnId.Text, out int toColumnId) && drawingColumnMarkers.ContainsKey(toColumnId) ? drawingColumnMarkers[toColumnId] : double.NaN;
                                double rightOffset = anchorTwoCell.ToMarker.ColumnOffset != null && int.TryParse(anchorTwoCell.ToMarker.ColumnOffset.Text, out int toMarkerColumnOffset) ? RoundNumber(toMarkerColumnOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                double bottom = configClone.ConvertSizes && anchorTwoCell.ToMarker != null && anchorTwoCell.ToMarker.RowId != null && int.TryParse(anchorTwoCell.ToMarker.RowId.Text, out int toRowId) && drawingRowMarkers.ContainsKey(toRowId) ? drawingRowMarkers[toRowId] : double.NaN;
                                double bottomOffset = anchorTwoCell.ToMarker.RowOffset != null && int.TryParse(anchorTwoCell.ToMarker.RowOffset.Text, out int toMarkerRowOffset) ? RoundNumber(toMarkerRowOffset / 914400.0 * 96, configClone.RoundingDigits) : 0;
                                drawing = GetDrawing(anchorTwoCell, !double.IsNaN(left) ? $"calc({GetInvariantNumber(left)}% + {GetInvariantNumber(leftOffset)}px)" : "0", !double.IsNaN(top) ? $"{GetInvariantNumber(RoundNumber(top + topOffset, configClone.RoundingDigits))}px" : "0", !double.IsNaN(left) && !double.IsNaN(right) ? $"calc({GetInvariantNumber(RoundNumber(right - left, configClone.RoundingDigits))}% + {GetInvariantNumber(RoundNumber(rightOffset - leftOffset, configClone.RoundingDigits))}px)" : "auto", !double.IsNaN(top) && !double.IsNaN(bottom) ? $"{GetInvariantNumber(RoundNumber(bottom + bottomOffset - top - topOffset, configClone.RoundingDigits))}px" : "auto", worksheetPart.DrawingsPart, themes, configClone);
                            }
                            if (!string.IsNullOrEmpty(drawing))
                            {
                                writer.Write($"\n{new string(' ', 8)}{drawing}");
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
                            writer.Write($"\n{new string(' ', 12)}{GetAttributesString(stylesheetCellFormats[stylesheetFormatIndex].Item1, false, 12)}");
                            writer.Write($"\n{new string(' ', 8)}}}{(stylesheetFormatIndex < stylesheetCellFormats.Length - 1 ? $"\n{new string(' ', 8)}" : string.Empty)}");
                        }
                    }
                    writer.Write($"\n{new string(' ', 4)}</style>");
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

        private static string GetInvariantNumber(double number)
        {
            return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetEscapedString(string value)
        {
            return System.Web.HttpUtility.HtmlEncode(value).Replace(" ", "&nbsp;");
        }

        private static Dictionary<string, string> GetJoinedAttributes(Dictionary<string, string> original, Dictionary<string, string> joining)
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

        private static string GetAttributesString(Dictionary<string, string> attributes, bool isAdditional, int indent)
        {
            if (attributes == null)
            {
                return string.Empty;
            }

            string result = string.Empty;
            foreach (KeyValuePair<string, string> pair in attributes)
            {
                result += $"{pair.Key}: {pair.Value};{(indent < 0 ? " " : $"\n{new string(' ', indent)}")}";
            }
            return !string.IsNullOrEmpty(result) ? (isAdditional ? $" {result.TrimEnd()}" : result.TrimEnd()) : string.Empty;
        }

        private static int GetColumnIndex(string cell)
        {
            int index = 0;
            Match match = regexLetters.Match(cell);
            if (match.Success)
            {
                string value = match.Value.ToUpper();
                for (int i = value.Length - 1, mulitplier = 1; i >= 0; i--, mulitplier *= 26)
                {
                    index += mulitplier * (value[i] - 64);
                }
            }
            return Math.Max(1, index);
        }

        private static int GetRowIndex(string cell)
        {
            Match match = regexNumbers.Match(cell);
            return match.Success && int.TryParse(match.Value, out int index) ? Math.Max(1, index) : 1;
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

        private static bool GetFormulaCondition(string value, IEnumerable<Formula> formulas, int formulasCount, Func<double[], bool> actionEvaluation)
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
                if (!double.TryParse(formula.Text, out double formulaDouble))
                {
                    continue;
                }

                index++;
                parameters[index] = formulaDouble;
                if (index > formulasCount)
                {
                    break;
                }
            }
            return index >= formulasCount && actionEvaluation.Invoke(parameters);
        }

        private static string GetFormattedNumber(string value, string format, ref Dictionary<string, (int, int, int, bool, int, int, int, int, bool, bool, string[])> formatsCalculated, Action<string[]> actionTestConditions)
        {
            if (string.IsNullOrEmpty(format) || format.ToLower() == "general")
            {
                return value;
            }

            bool isValueNumber = double.TryParse(value, out double valueNumber);
            if (!isValueNumber && !format.Contains("@"))
            {
                return value;
            }
            else if (isValueNumber)
            {
                valueNumber = Math.Abs(valueNumber);
            }

            int infoValue = value.Length;
            bool isFormatCalculated = formatsCalculated.ContainsKey(format);
            (int, int, int, bool, int, int, int, int, bool, bool, string[]) infoFormat = isFormatCalculated ? formatsCalculated[format] : (format.Length, format.Length, format.Length, false, -1, -1, -1, -1, false, false, null);
            List<string> infoFormatConditions = null;
            (bool, string) valueScientific = (true, "0");
            bool isFormattingScientific = false;
            (string, string) valueFraction = (string.Empty, string.Empty);
            bool isFormattingFraction = false;
            void actionUpdateValue()
            {
                value = isValueNumber ? valueNumber.ToString(stringFormatNumber) : value;
                infoValue = value.IndexOf('.');
                infoValue = infoValue < 0 ? value.Length : infoValue;
            }
            actionUpdateValue();

            int indexValue = 0;
            int indexFormat = 0;
            bool isIncreasing = true;
            void actionUpdateInfo()
            {
                if (isValueNumber && infoFormat.Item4)
                {
                    valueNumber *= 100;
                    actionUpdateValue();
                }
                if (isValueNumber && infoFormat.Item5 >= 0)
                {
                    if (infoValue > 1)
                    {
                        valueScientific = (true, (infoValue - 1).ToString());
                        valueNumber /= Math.Pow(10, infoValue - 1);
                        actionUpdateValue();
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
                            actionUpdateValue();
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
                    if (valueAbsolute == 0 || valueAbsolute < maxError)
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
                        void actionFindNewValue(int indexNumerator, int indexDenominator, int incrementNumerator, int incrementDenominator, Func<int, int, bool> actionEvaluation)
                        {
                            fractionParts[indexNumerator] += incrementNumerator;
                            fractionParts[indexDenominator] += incrementDenominator;
                            if (!actionEvaluation.Invoke(fractionParts[indexNumerator], fractionParts[indexDenominator]))
                            {
                                return;
                            }

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
                                fractionNumerator = valueNumber >= 0 ? valueFloor * middleDenominator + middleNumerator : -(valueFloor * middleDenominator + middleNumerator);
                                fractionDenominator = middleDenominator;
                                break;
                            }
                        }
                    }
                    valueFraction = (fractionNumerator.ToString(), fractionDenominator.ToString());
                    actionUpdateValue();
                }
                indexValue = infoValue;
                indexFormat = infoFormat.Item2 - 1;
                isIncreasing = false;
            }
            if (isFormatCalculated)
            {
                actionUpdateInfo();
            }

            string result = string.Empty;
            string resultFormatted = string.Empty;
            string actionReverseString(string original)
            {
                string reversed = string.Empty;
                for (int i = original.Length - 1; i >= 0; i--)
                {
                    reversed += original[i];
                }
                return reversed;
            }
            while (indexFormat < format.Length || !isFormatCalculated)
            {
                if (isFormatCalculated && !isIncreasing && isFormattingScientific && infoFormat.Item5 >= 0 && indexFormat > 0 && format[indexFormat - 1] == 'E' && (format[indexFormat] == '+' || format[indexFormat] == '-'))
                {
                    result = resultFormatted + actionReverseString(result);
                    resultFormatted = string.Empty;
                    indexFormat = infoFormat.Item5 + 1;
                    isIncreasing = true;
                    isFormattingScientific = false;
                    continue;
                }
                else if (isFormatCalculated && !isIncreasing && isFormattingFraction && infoFormat.Item7 >= 0 && indexFormat < infoFormat.Item6)
                {
                    result = resultFormatted + actionReverseString(result);
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
                    infoFormat.Item11 = infoFormatConditions != null ? infoFormatConditions.ToArray() : infoFormat.Item11;
                    formatsCalculated[format] = infoFormat;
                    isFormatCalculated = true;
                    actionUpdateInfo();
                    continue;
                }
                else if (indexFormat < 0)
                {
                    result = actionReverseString(result);
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
                    string condition = string.Empty;
                    infoFormatConditions = !isFormatCalculated && infoFormatConditions == null ? new List<string>() : infoFormatConditions;
                    do
                    {
                        indexFormat += isIncreasing ? 1 : -1;
                        condition += !isFormatCalculated ? format[indexFormat].ToString() : string.Empty;
                    } while (isIncreasing ? indexFormat + 1 < format.Length && format[indexFormat + 1] != ']' : indexFormat > 0 && format[indexFormat - 1] != '[');
                    indexFormat += isIncreasing ? 2 : -2;
                    if (!isFormatCalculated)
                    {
                        infoFormatConditions.Add(condition.Trim());
                    }
                    continue;
                }
                else if (formatChar == '"' && (isIncreasing ? indexFormat + 1 < format.Length : indexFormat > 0))
                {
                    do
                    {
                        indexFormat += isIncreasing ? 1 : -1;
                        result += isFormatCalculated ? format[indexFormat].ToString() : string.Empty;
                    }
                    while (isIncreasing ? indexFormat + 1 < format.Length && format[indexFormat + 1] != '"' : indexFormat > 0 && format[indexFormat - 1] != '"');
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
                        infoFormat.Item9 |= indexFormat > infoFormat.Item2 && (formatChar == '0' || formatChar == '?') && infoFormat.Item5 < 0;
                    }
                    else if (formatChar == ' ')
                    {
                        infoFormat.Item10 |= indexFormat > infoFormat.Item1;
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
                        result += isIncreasing ? value : actionReverseString(value);
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
                                return GetFormattedNumber((valueNumber + (10 - next) / Math.Pow(10, indexValue + 1 - infoValue)).ToString(stringFormatNumber), format, ref formatsCalculated, actionTestConditions);
                            }

                            result += (!isFormattingScientific ? (!isFormattingFraction ? value : (indexFormat > infoFormat.Item7 ? valueFraction.Item2 : valueFraction.Item1)) : valueScientific.Item2)[indexValue].ToString();
                            if (!isIncreasing && (!isFormattingScientific ? (!isFormattingFraction ? indexFormat <= infoFormat.Item1 : indexFormat <= infoFormat.Item6) : indexFormat - 2 >= 0 && format[indexFormat - 2] == 'E' && (format[indexFormat - 1] == '+' || format[indexFormat - 1] == '-')))
                            {
                                result += actionReverseString((!isFormattingScientific ? (!isFormattingFraction ? value : valueFraction.Item1) : valueScientific.Item2).Substring(0, indexValue));
                            }
                            else if (isIncreasing && isFormattingFraction && indexFormat >= infoFormat.Item8 && indexValue + 1 < valueFraction.Item2.Length)
                            {
                                result += valueFraction.Item2.Substring(indexValue + 1, valueFraction.Item2.Length - indexValue - 1);
                            }
                        }
                        else
                        {
                            result += formatChar == '0' && (!isFormattingFraction || indexFormat <= infoFormat.Item7 || indexFormat > infoFormat.Item8) ? "0" : (formatChar == '?' ? " " : string.Empty);
                        }
                    }
                    else
                    {
                        result += formatChar;
                    }
                }
                indexFormat += isIncreasing ? 1 : -1;
            }
            if (infoFormat.Item11 != null)
            {
                actionTestConditions.Invoke(infoFormat.Item11);
            }
            return result;
        }

        private static void GetFont(OpenXmlElement font, ref Dictionary<string, string> styles, ref string valueContainer, DocumentFormat.OpenXml.Drawing.Color2Type[] themes, ConverterConfig config)
        {
            styles = styles ?? new Dictionary<string, string>();
            if (font == null)
            {
                return;
            }

            bool isTextDecoraionSet = false;
            string stylesTextDecoraion = string.Empty;
            if (font is DocumentFormat.OpenXml.Drawing.TextCharacterPropertiesType fontPropertiesType)
            {
                if (fontPropertiesType.FontSize != null && fontPropertiesType.FontSize.HasValue)
                {
                    styles["font-size"] = $"{GetInvariantNumber(RoundNumber(fontPropertiesType.FontSize.Value / 7200.0 * 96, config.RoundingDigits))}px";
                }
                if (fontPropertiesType.Bold != null)
                {
                    styles["font-weight"] = !fontPropertiesType.Bold.HasValue || fontPropertiesType.Bold.Value ? "bold" : "normal";
                }
                if (fontPropertiesType.Italic != null)
                {
                    styles["font-style"] = !fontPropertiesType.Italic.HasValue || fontPropertiesType.Italic.Value ? "italic" : "normal";
                }
                if (fontPropertiesType.Strike != null)
                {
                    if (!fontPropertiesType.Strike.HasValue || fontPropertiesType.Strike.Value == DocumentFormat.OpenXml.Drawing.TextStrikeValues.SingleStrike)
                    {
                        stylesTextDecoraion += " line-through";
                    }
                    else if (fontPropertiesType.Strike.Value == DocumentFormat.OpenXml.Drawing.TextStrikeValues.DoubleStrike)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: line-through double;\">{{0}}</span>");
                    }
                    isTextDecoraionSet = true;
                }
                if (fontPropertiesType.Underline != null)
                {
                    if (!fontPropertiesType.Underline.HasValue || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Words)
                    {
                        stylesTextDecoraion += " underline";
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Heavy)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline 4px;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Double)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline double;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dash || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLong || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDash)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline dashed;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashHeavy || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLongHeavy || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDashHeavy)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline dashed 4px;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dotted || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDash)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline dotted;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.HeavyDotted || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDashHeavy)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline dotted 4px;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Wavy)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline wavy;\">{{0}}</span>");
                    }
                    else if (fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyDouble || fontPropertiesType.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyHeavy)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline wavy 4px;\">{{0}}</span>");
                    }
                    isTextDecoraionSet = true;
                }
                if (fontPropertiesType.Spacing != null && fontPropertiesType.Spacing.HasValue)
                {
                    styles["letter-spacing"] = $"{GetInvariantNumber(RoundNumber(fontPropertiesType.Spacing.Value / 7200.0 * 96, config.RoundingDigits))}px";
                }
                if (fontPropertiesType.Capital != null)
                {
                    styles["text-transform"] = !fontPropertiesType.Capital.HasValue || fontPropertiesType.Capital.Value == DocumentFormat.OpenXml.Drawing.TextCapsValues.All ? "uppercase" : (fontPropertiesType.Capital.Value == DocumentFormat.OpenXml.Drawing.TextCapsValues.Small ? "lowercase" : "none");
                }
            }
            foreach (OpenXmlElement fontElement in font.Elements())
            {
                if (fontElement is RunFont fontNameRun && fontNameRun.Val != null && fontNameRun.Val.HasValue)
                {
                    styles["font-family"] = fontNameRun.Val.Value;
                }
                else if (fontElement is FontName fontName && fontName.Val != null && fontName.Val.HasValue)
                {
                    styles["font-family"] = fontName.Val.Value;
                }
                else if (fontElement is Color fontFill && GetColor(fontFill, out string fontColor, themes, config))
                {
                    styles["color"] = fontColor;
                }
                else if (fontElement is FontSize fontSize && fontSize.Val != null && fontSize.Val.HasValue)
                {
                    styles["font-size"] = $"{GetInvariantNumber(RoundNumber(fontSize.Val.Value / 72 * 96, config.RoundingDigits))}px";
                }
                else if (fontElement is Bold fontBold)
                {
                    styles["font-weight"] = fontBold.Val == null || !fontBold.Val.HasValue || fontBold.Val.Value ? "bold" : "normal";
                }
                else if (fontElement is Italic fontItalic)
                {
                    styles["font-style"] = fontItalic.Val == null || !fontItalic.Val.HasValue || fontItalic.Val.Value ? "italic" : "normal";
                }
                else if (fontElement is Strike fontStrike)
                {
                    if (fontStrike.Val == null || !fontStrike.Val.HasValue || fontStrike.Val.Value)
                    {
                        stylesTextDecoraion += " line-through";
                    }
                    isTextDecoraionSet = true;
                }
                else if (fontElement is Underline underline)
                {
                    if (underline.Val == null || !underline.Val.HasValue || underline.Val.Value == UnderlineValues.Single || underline.Val.Value == UnderlineValues.SingleAccounting)
                    {
                        stylesTextDecoraion += " underline";
                    }
                    else if (underline.Val.Value != UnderlineValues.None)
                    {
                        valueContainer = valueContainer.Replace("{0}", $"<span style=\"text-decoration: underline double;\">{{0}}</span>");
                    }
                    isTextDecoraionSet = true;
                }
                else if (fontElement is VerticalTextAlignment fontVerticalAlignment && fontVerticalAlignment.Val != null)
                {
                    styles["vertical-align"] = fontVerticalAlignment.Val.Value == VerticalAlignmentRunValues.Subscript ? "sub" : (fontVerticalAlignment.Val.Value == VerticalAlignmentRunValues.Superscript ? "super" : "baseline");
                }
                else if (fontElement is Extend fontExtend)
                {
                    styles["font-stretch"] = fontExtend.Val == null || !fontExtend.Val.HasValue || fontExtend.Val.Value ? "expanded" : "normal";
                }
                else if (fontElement is Condense fontCondense)
                {
                    styles["font-stretch"] = fontCondense.Val == null || !fontCondense.Val.HasValue || fontCondense.Val.Value ? "condensed" : "normal";
                }
                else if (fontElement is DocumentFormat.OpenXml.Drawing.NoFill)
                {
                    styles["color"] = "transparent";
                }
                else if (fontElement is DocumentFormat.OpenXml.Drawing.SolidFill fontDrawingFill && GetColor(fontDrawingFill, out string fontDrawingColor, themes, config))
                {
                    styles["color"] = fontDrawingColor;
                }
                else if (fontElement is DocumentFormat.OpenXml.Drawing.TextFontType fontDrawingName && fontDrawingName.Typeface != null && fontDrawingName.Typeface.HasValue)
                {
                    styles["font-family"] = fontDrawingName.Typeface.Value;
                }
                else if (fontElement is DocumentFormat.OpenXml.Drawing.RightToLeft fontDrawingRightToLeft)
                {
                    styles["direction"] = fontDrawingRightToLeft.Val == null || !fontDrawingRightToLeft.Val.HasValue || fontDrawingRightToLeft.Val.Value ? "rtl" : "ltr";
                }
            }
            if (isTextDecoraionSet)
            {
                styles["text-decoration"] = !string.IsNullOrEmpty(stylesTextDecoraion) ? stylesTextDecoraion.TrimStart() : "none";
            }
        }

        private static Dictionary<string, string> GetCellFormat(Fill fill, Font font, Border border, Alignment alignment, ref string valueContainer, DocumentFormat.OpenXml.Drawing.Color2Type[] themes, ConverterConfig config)
        {
            Dictionary<string, string> styles = new Dictionary<string, string>();
            if (fill != null)
            {
                if (fill.PatternFill != null && (fill.PatternFill.PatternType == null || !fill.PatternFill.PatternType.HasValue || fill.PatternFill.PatternType.Value != PatternValues.None))
                {
                    if (fill.PatternFill.BackgroundColor != null && GetColor(fill.PatternFill.BackgroundColor, out string fillColorBackground, themes, config))
                    {
                        styles["background"] = fillColorBackground;
                    }
                    string fillColorForeground = fill.PatternFill.ForegroundColor != null && GetColor(fill.PatternFill.ForegroundColor, out fillColorForeground, themes, config) ? fillColorForeground : "black";
                    if (fill.PatternFill.PatternType != null && fill.PatternFill.PatternType.HasValue)
                    {
                        if (fill.PatternFill.PatternType.Value == PatternValues.DarkGray)
                        {
                            styles["background-image"] = $"radial-gradient(circle at 1px 1px, {fillColorForeground} 0.5px, transparent 0), radial-gradient(circle at 2.6px 2.6px, {fillColorForeground} 0.5px, transparent 0)";
                            styles["background-size"] = "3.2px 3.2px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.MediumGray)
                        {
                            styles["background-image"] = $"radial-gradient(circle at 1px 1px, {fillColorForeground} 0.5px, transparent 0), radial-gradient(circle at 2.8px 2.8px, {fillColorForeground} 0.5px, transparent 0)";
                            styles["background-size"] = "3.6px 3.6px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.LightGray)
                        {
                            styles["background-image"] = $"radial-gradient(circle at 1px 1px, {fillColorForeground} 0.5px, transparent 0), radial-gradient(circle at 3px 3px, {fillColorForeground} 0.5px, transparent 0)";
                            styles["background-size"] = "4px 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.Gray125)
                        {
                            styles["background-image"] = $"radial-gradient(circle at 1px 1px, {fillColorForeground} 0.5px, transparent 0), radial-gradient(circle at 4px 4px, {fillColorForeground} 0.5px, transparent 0)";
                            styles["background-size"] = "6px 6px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.Gray0625)
                        {
                            styles["background-image"] = $"radial-gradient(circle at 1px 1px, {fillColorForeground} 0.5px, transparent 0), radial-gradient(circle at 5.5px 5.5px, {fillColorForeground} 0.5px, transparent 0)";
                            styles["background-size"] = "9px 9px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.DarkHorizontal)
                        {
                            styles["background-image"] = $"linear-gradient(0deg, {fillColorForeground} 1.5px, transparent 0)";
                            styles["background-size"] = "100% 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.LightHorizontal)
                        {
                            styles["background-image"] = $"linear-gradient(0deg, {fillColorForeground} 1px, transparent 0)";
                            styles["background-size"] = "100% 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.DarkVertical)
                        {
                            styles["background-image"] = $"linear-gradient(90deg, {fillColorForeground} 1.5px, transparent 0)";
                            styles["background-size"] = "4px 100%";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.LightVertical)
                        {
                            styles["background-image"] = $"linear-gradient(90deg, {fillColorForeground} 1px, transparent 0)";
                            styles["background-size"] = "4px 100%";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.DarkDown)
                        {
                            styles["background-image"] = $"linear-gradient(45deg, {fillColorForeground} 25%, transparent 25% 50%, {fillColorForeground} 50% 75%, transparent 75%)";
                            styles["background-size"] = "4px 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.LightDown)
                        {
                            styles["background-image"] = $"linear-gradient(45deg, {fillColorForeground} 10%, transparent 10% 50%, {fillColorForeground} 50% 60%, transparent 60%)";
                            styles["background-size"] = "4px 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.DarkUp)
                        {
                            styles["background-image"] = $"linear-gradient(-45deg, {fillColorForeground} 25%, transparent 25% 50%, {fillColorForeground} 50% 75%, transparent 75%)";
                            styles["background-size"] = "4px 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.DarkGrid)
                        {
                            styles["background-image"] = $"linear-gradient(45deg, {fillColorForeground} 25%, transparent 25% 75%, {fillColorForeground} 75%), linear-gradient(45deg, {fillColorForeground} 25%, transparent 25% 75%, {fillColorForeground} 75%)";
                            styles["background-position"] = "0 0, 2.5px 2.5px";
                            styles["background-size"] = "5px 5px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.LightGrid)
                        {
                            styles["background-image"] = $"linear-gradient(90deg, {fillColorForeground} 1px, transparent 0), linear-gradient(0deg, {fillColorForeground} 1px, transparent 0)";
                            styles["background-size"] = "4px 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.DarkTrellis)
                        {
                            styles["background-image"] = $"linear-gradient(45deg, {fillColorForeground} 15%, transparent 15% 50%, {fillColorForeground} 50% 65%, transparent 65%), linear-gradient(-45deg, {fillColorForeground} 15%, transparent 15% 50%, {fillColorForeground} 50% 65%, transparent 65%)";
                            styles["background-size"] = "4px 4px";
                        }
                        else if (fill.PatternFill.PatternType.Value == PatternValues.LightTrellis)
                        {
                            styles["background-image"] = $"linear-gradient(45deg, {fillColorForeground} 10%, transparent 10% 50%, {fillColorForeground} 50% 60%, transparent 60%), linear-gradient(-45deg, {fillColorForeground} 10%, transparent 10% 50%, {fillColorForeground} 50% 60%, transparent 60%)";
                            styles["background-size"] = "4px 4px";
                        }
                        else
                        {
                            styles["background"] = fillColorForeground;
                        }
                    }
                    else
                    {
                        styles["background"] = fillColorForeground;
                    }
                }
                else if (fill.GradientFill != null)
                {
                    if (fill.GradientFill.Type == null || !fill.GradientFill.Type.HasValue || fill.GradientFill.Type.Value == GradientValues.Linear)
                    {
                        string fillColor = $"linear-gradient({(fill.GradientFill.Degree != null && fill.GradientFill.Degree.HasValue ? GetInvariantNumber(RoundNumber(((fill.GradientFill.Degree.Value + 90) % 360 + 360) % 360, config.RoundingDigits)) : "90")}deg";
                        foreach (GradientStop gradient in fill.GradientFill.Elements<GradientStop>())
                        {
                            if (gradient.Color != null && GetColor(gradient.Color, out string gradientColor, themes, config))
                            {
                                fillColor += $", {gradientColor}{(gradient.Position != null && gradient.Position.HasValue ? $" {GetInvariantNumber(RoundNumber(gradient.Position.Value * 100, config.RoundingDigits))}%" : string.Empty)}";
                            }
                        }
                        styles["background"] = $"{fillColor})";
                    }
                    else
                    {
                        double gradientLeft = fill.GradientFill.Left != null && fill.GradientFill.Left.HasValue ? fill.GradientFill.Left.Value : 0;
                        double gradientTop = fill.GradientFill.Top != null && fill.GradientFill.Top.HasValue ? fill.GradientFill.Top.Value : 0;
                        double gradientRight = fill.GradientFill.Right != null && fill.GradientFill.Right.HasValue ? fill.GradientFill.Right.Value : 0;
                        double gradientBottom = fill.GradientFill.Bottom != null && fill.GradientFill.Bottom.HasValue ? fill.GradientFill.Bottom.Value : 0;
                        double gradientRadius = ((gradientLeft + gradientRight) / 2 + (gradientTop + gradientBottom) / 2 - gradientLeft - gradientTop) / 2;
                        string fillColor = $"radial-gradient(circle at {GetInvariantNumber(RoundNumber((gradientLeft + gradientRight) / 2 * 100, config.RoundingDigits))}% {GetInvariantNumber(RoundNumber((gradientTop + gradientBottom) / 2 * 100, config.RoundingDigits))}%";
                        foreach (GradientStop gradient in fill.GradientFill.Elements<GradientStop>())
                        {
                            if (gradient.Color != null && GetColor(gradient.Color, out string gradientColor, themes, config))
                            {
                                fillColor += $", {gradientColor}{(gradient.Position != null && gradient.Position.HasValue ? $" {GetInvariantNumber(RoundNumber((gradientRadius + gradient.Position.Value * (1 - gradientRadius)) * 100, config.RoundingDigits))}%" : string.Empty)}";
                            }
                        }
                        styles["background"] = $"{fillColor})";
                    }
                }
            }
            if (font != null)
            {
                GetFont(font, ref styles, ref valueContainer, themes, config);
            }
            if (border != null)
            {
                bool actionGetBorder(BorderPropertiesType borderProperties, out string attribute)
                {
                    if (borderProperties == null)
                    {
                        attribute = string.Empty;
                        return false;
                    }

                    string result = string.Empty;
                    if (borderProperties.Style != null && borderProperties.Style.HasValue)
                    {
                        if (borderProperties.Style.Value == BorderStyleValues.None)
                        {
                            attribute = string.Empty;
                            return false;
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.Thick)
                        {
                            result += " thick solid";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.Medium)
                        {
                            result += " medium solid";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.MediumDashed || borderProperties.Style.Value == BorderStyleValues.MediumDashDot)
                        {
                            result += " medium dashed";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.MediumDashDotDot)
                        {
                            result += " medium dotted";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.Double)
                        {
                            result += " medium double";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.Thin)
                        {
                            result += " thin solid";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.Dashed || borderProperties.Style.Value == BorderStyleValues.DashDot || borderProperties.Style.Value == BorderStyleValues.SlantDashDot)
                        {
                            result += " thin dashed";
                        }
                        else if (borderProperties.Style.Value == BorderStyleValues.DashDotDot || borderProperties.Style.Value == BorderStyleValues.Hair)
                        {
                            result += " thin dotted";
                        }
                    }
                    if (borderProperties.Color != null && GetColor(borderProperties.Color, out string borderColor, themes, config))
                    {
                        result += $" {borderColor}";
                    }
                    attribute = result.TrimStart();
                    return !string.IsNullOrEmpty(attribute);
                }

                if (actionGetBorder(border.TopBorder, out string borderTop))
                {
                    styles["border-top"] = borderTop;
                }
                if (actionGetBorder(border.RightBorder, out string borderRight))
                {
                    styles["border-right"] = borderRight;
                }
                if (actionGetBorder(border.BottomBorder, out string borderBottom))
                {
                    styles["border-bottom"] = borderBottom;
                }
                if (actionGetBorder(border.LeftBorder, out string borderLeft))
                {
                    styles["border-left"] = borderLeft;
                }
            }
            if (alignment != null)
            {
                if (alignment.Horizontal != null && alignment.Horizontal.HasValue && alignment.Horizontal.Value != HorizontalAlignmentValues.General)
                {
                    styles["text-align"] = alignment.Horizontal.Value == HorizontalAlignmentValues.Left ? "left" : (alignment.Horizontal.Value == HorizontalAlignmentValues.Right ? "right" : (alignment.Horizontal.Value == HorizontalAlignmentValues.Center || alignment.Horizontal.Value == HorizontalAlignmentValues.CenterContinuous ? "center" : "justify"));
                }
                if (alignment.Vertical != null && alignment.Vertical.HasValue)
                {
                    styles["vertical-align"] = alignment.Vertical.Value == VerticalAlignmentValues.Bottom ? "bottom" : (alignment.Vertical.Value == VerticalAlignmentValues.Top ? "top" : "middle");
                }
                if (alignment.WrapText != null && (!alignment.WrapText.HasValue || alignment.WrapText.Value))
                {
                    styles["overflow-wrap"] = "break-word";
                    styles["white-space"] = "normal";
                }
                if (alignment.TextRotation != null && alignment.TextRotation.HasValue && alignment.TextRotation.Value != 0)
                {
                    valueContainer = valueContainer.Replace("{0}", $"<div style=\"width: fit-content; transform: rotate(-{GetInvariantNumber(RoundNumber(alignment.TextRotation.Value, config.RoundingDigits))}deg);\">{{0}}</div>");
                }
            }
            return styles.Count > 0 ? styles : null;
        }

        private static string GetDrawing(OpenXmlElement drawing, string left, string top, string width, string height, DrawingsPart drawingsPart, DocumentFormat.OpenXml.Drawing.Color2Type[] themes, ConverterConfig config)
        {
            if (drawing == null)
            {
                return string.Empty;
            }

            string element = string.Empty;
            double elementPaddingTop = 0;
            double elementPaddingRight = 0;
            double elementPaddingBottom = 0;
            double elementPaddingLeft = 0;
            Dictionary<string, string> elementStyles = new Dictionary<string, string>();
            DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties elementShapeProperties = null;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeStyle elementShapeStyle = null;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties elementNonVisualProperties = null;
            foreach (OpenXmlElement drawingElement in drawing.Elements())
            {
                if (drawingElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture && config.ConvertPictures && picture.BlipFill != null && picture.BlipFill.Blip != null && picture.BlipFill.Blip.Embed != null && picture.BlipFill.Blip.Embed.HasValue && drawingsPart != null && drawingsPart.GetPartById(picture.BlipFill.Blip.Embed.Value) is ImagePart imagePart)
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
                    element = $"<img loading=\"lazy\" decoding=\"async\" style=\"{{0}}\" src=\"data:{imagePart.ContentType};base64,{base64}\"{description} />";
                    elementShapeProperties = picture.ShapeProperties;
                    elementShapeStyle = picture.ShapeStyle;
                    elementNonVisualProperties = picture.NonVisualPictureProperties.NonVisualDrawingProperties;
                    break;
                }
                else if (drawingElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape && config.ConvertShapes)
                {
                    string shapeValue = !config.ConvertStyles && shape.TextBody != null ? shape.TextBody.InnerText : string.Empty;
                    elementStyles["overflow-wrap"] = "break-word";
                    elementStyles["white-space"] = "normal";
                    elementStyles["overflow-x"] = "hidden";
                    elementStyles["overflow-y"] = "hidden";
                    elementPaddingTop = 0.13 / 2.54 * 96;
                    elementPaddingRight = 0.25 / 2.54 * 96;
                    elementPaddingBottom = 0.13 / 2.54 * 96;
                    elementPaddingLeft = 0.25 / 2.54 * 96;
                    if (shape.TextBody != null)
                    {
                        foreach (OpenXmlElement textBodyElement in shape.TextBody.Elements())
                        {
                            if (textBodyElement is DocumentFormat.OpenXml.Drawing.Paragraph paragraph && config.ConvertStyles)
                            {
                                string paragraphValue = string.Empty;
                                Dictionary<string, string> paragraphStyles = new Dictionary<string, string>();
                                string paragraphValueContainer = "{0}";
                                foreach (OpenXmlElement paragraphElement in paragraph.Elements())
                                {
                                    if (paragraphElement is DocumentFormat.OpenXml.Drawing.Text text)
                                    {
                                        paragraphValue += GetEscapedString(text.Text);
                                    }
                                    else if (paragraphElement is DocumentFormat.OpenXml.Drawing.Run run && run.Text != null)
                                    {
                                        string runValueContainer = "{0}";
                                        Dictionary<string, string> runStyles = null;
                                        if (run.RunProperties is DocumentFormat.OpenXml.Drawing.RunProperties runProperties)
                                        {
                                            GetFont(runProperties, ref runStyles, ref runValueContainer, themes, config);
                                        }
                                        paragraphValue += $"<span{(runStyles != null && runStyles.Count > 0 ? $" style=\"{GetAttributesString(runStyles, false, -1)}\"" : string.Empty)}>{runValueContainer.Replace("{0}", GetEscapedString(run.Text.Text))}</span>";
                                    }
                                    else if (paragraphElement is DocumentFormat.OpenXml.Drawing.ParagraphProperties paragraphProperties)
                                    {
                                        if (paragraphProperties.Alignment != null && paragraphProperties.Alignment.HasValue && paragraphProperties.Alignment.Value != DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left)
                                        {
                                            paragraphStyles["text-align"] = paragraphProperties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right ? "right" : (paragraphProperties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center ? "center" : "justify");
                                        }
                                        if (paragraphProperties.LeftMargin != null && paragraphProperties.LeftMargin.HasValue && paragraphProperties.LeftMargin.Value != 0)
                                        {
                                            paragraphStyles["margin-left"] = $"{GetInvariantNumber(RoundNumber(paragraphProperties.LeftMargin.Value / 914400.0 * 96, config.RoundingDigits))}px";
                                        }
                                        if (paragraphProperties.RightMargin != null && paragraphProperties.RightMargin.HasValue && paragraphProperties.RightMargin.Value != 0)
                                        {
                                            paragraphStyles["margin-right"] = $"{GetInvariantNumber(RoundNumber(paragraphProperties.RightMargin.Value / 914400.0 * 96, config.RoundingDigits))}px";
                                        }
                                        if (paragraphProperties.Indent != null && paragraphProperties.Indent.HasValue && paragraphProperties.Indent.Value != 0)
                                        {
                                            paragraphStyles["text-indent"] = $"{GetInvariantNumber(RoundNumber(paragraphProperties.Indent.Value / 914400.0 * 96, config.RoundingDigits))}px";
                                        }
                                        if (paragraphProperties.DefaultTabSize != null && paragraphProperties.DefaultTabSize.HasValue)
                                        {
                                            paragraphStyles["tab-size"] = $"{GetInvariantNumber(RoundNumber(paragraphProperties.DefaultTabSize.Value / 914400.0 * 96, config.RoundingDigits))}px";
                                        }
                                        if (paragraphProperties.RightToLeft != null && paragraphProperties.RightToLeft.HasValue && paragraphProperties.RightToLeft.Value)
                                        {
                                            paragraphStyles["direction"] = "rtl";
                                        }
                                        foreach (OpenXmlElement propertiesElement in paragraphProperties.Elements())
                                        {
                                            if (propertiesElement is DocumentFormat.OpenXml.Drawing.DefaultRunProperties paragraphRunProperties)
                                            {
                                                string valueContainer = "{0}";
                                                GetFont(paragraphRunProperties, ref paragraphStyles, ref valueContainer, themes, config);
                                                paragraphValueContainer = valueContainer.Replace("{0}", paragraphValueContainer);
                                            }
                                            else if (propertiesElement is DocumentFormat.OpenXml.Drawing.LineSpacing paragraphLineSpacing)
                                            {
                                                if (paragraphLineSpacing.SpacingPercent != null && paragraphLineSpacing.SpacingPercent.Val != null && paragraphLineSpacing.SpacingPercent.Val.HasValue)
                                                {
                                                    paragraphStyles["line-height"] = GetInvariantNumber(RoundNumber(paragraphLineSpacing.SpacingPercent.Val.Value / 100000.0, config.RoundingDigits));
                                                }
                                                else if (paragraphLineSpacing.SpacingPoints != null && paragraphLineSpacing.SpacingPoints.Val != null && paragraphLineSpacing.SpacingPoints.Val.HasValue)
                                                {
                                                    paragraphStyles["line-height"] = $"{GetInvariantNumber(RoundNumber(paragraphLineSpacing.SpacingPoints.Val.Value / 7200.0 * 96, config.RoundingDigits))}px";
                                                }
                                            }
                                        }
                                    }
                                    else if (paragraphElement is DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties endParagraphRunProperties)
                                    {
                                        string valueContainer = "{0}";
                                        GetFont(endParagraphRunProperties, ref elementStyles, ref valueContainer, themes, config);
                                        paragraphValueContainer = valueContainer.Replace("{0}", paragraphValueContainer);
                                    }
                                }
                                shapeValue += $"<div style=\"{GetAttributesString(paragraphStyles, false, -1)}\">{paragraphValue}</div>";
                            }
                            else if (textBodyElement is DocumentFormat.OpenXml.Drawing.BodyProperties bodyProperties)
                            {
                                elementPaddingTop = bodyProperties.TopInset != null && bodyProperties.TopInset.HasValue ? bodyProperties.TopInset.Value / 914400.0 * 96 : elementPaddingTop;
                                elementPaddingRight = bodyProperties.RightInset != null && bodyProperties.RightInset.HasValue ? bodyProperties.RightInset.Value / 914400.0 * 96 : elementPaddingRight;
                                elementPaddingBottom = bodyProperties.BottomInset != null && bodyProperties.BottomInset.HasValue ? bodyProperties.BottomInset.Value / 914400.0 * 96 : elementPaddingBottom;
                                elementPaddingLeft = bodyProperties.LeftInset != null && bodyProperties.LeftInset.HasValue ? bodyProperties.LeftInset.Value / 914400.0 * 96 : elementPaddingLeft;
                                if (bodyProperties.Anchor != null && bodyProperties.Anchor.HasValue && bodyProperties.Anchor.Value != DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Top)
                                {
                                    elementStyles["align-content"] = bodyProperties.Anchor.Value == DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Center ? "center" : "end";
                                }
                                if (bodyProperties.AnchorCenter != null && bodyProperties.AnchorCenter.HasValue && bodyProperties.AnchorCenter.Value)
                                {
                                    elementStyles["text-align"] = "center";
                                }
                                if (bodyProperties.Wrap != null && bodyProperties.Wrap.HasValue && bodyProperties.Wrap.Value == DocumentFormat.OpenXml.Drawing.TextWrappingValues.None)
                                {
                                    elementStyles.Remove("overflow-wrap");
                                    elementStyles.Remove("white-space");
                                }
                                if (bodyProperties.HorizontalOverflow != null && bodyProperties.HorizontalOverflow.HasValue && bodyProperties.HorizontalOverflow.Value == DocumentFormat.OpenXml.Drawing.TextHorizontalOverflowValues.Overflow)
                                {
                                    elementStyles.Remove("overflow-x");
                                }
                                if (bodyProperties.VerticalOverflow != null && bodyProperties.VerticalOverflow.HasValue && bodyProperties.VerticalOverflow.Value == DocumentFormat.OpenXml.Drawing.TextVerticalOverflowValues.Overflow)
                                {
                                    elementStyles.Remove("overflow-y");
                                }
                                if (bodyProperties.Vertical != null && bodyProperties.Vertical.HasValue && bodyProperties.Vertical.Value != DocumentFormat.OpenXml.Drawing.TextVerticalValues.Horizontal)
                                {
                                    if (bodyProperties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.Vertical270)
                                    {
                                        elementStyles["writing-mode"] = "vertical-rl";
                                    }
                                    else
                                    {
                                        shapeValue = $"<div style=\"width: fit-content; writing-mode: vertical-lr; transform: rotate(180deg);\">{shapeValue}</div>";
                                    }
                                }
                                if (bodyProperties.Rotation != null && bodyProperties.Rotation.HasValue)
                                {
                                    shapeValue = $"<div style=\"width: fit-content; transform: rotate({GetInvariantNumber(RoundNumber(bodyProperties.Rotation.Value / 60000.0, config.RoundingDigits))}deg);\">{shapeValue}</div>";
                                }
                            }
                        }
                    }
                    element = $"<div style=\"box-sizing: border-box; {{0}}\">{shapeValue}</div>";
                    elementShapeProperties = shape.ShapeProperties;
                    elementShapeStyle = shape.ShapeStyle;
                    elementNonVisualProperties = shape.NonVisualShapeProperties.NonVisualDrawingProperties;
                    break;
                }
            }
            if (string.IsNullOrEmpty(element))
            {
                return string.Empty;
            }
            string widthActual = width;
            string heightActual = height;
            bool isFillHandled = false;
            bool isOutlineHandled = false;
            if (elementShapeProperties != null)
            {
                foreach (OpenXmlElement propertiesElement in elementShapeProperties.Elements())
                {
                    if (propertiesElement is DocumentFormat.OpenXml.Drawing.Transform2D propertiesTransform)
                    {
                        string elementStylesTransforms = string.Empty;
                        if (propertiesTransform.Offset != null)
                        {
                            if (left == "0" && propertiesTransform.Offset.X != null && propertiesTransform.Offset.X.HasValue && propertiesTransform.Offset.X.Value != 0)
                            {
                                elementStylesTransforms += $" translateX({GetInvariantNumber(RoundNumber(propertiesTransform.Offset.X.Value / 914400.0 * 96, config.RoundingDigits))}px)";
                            }
                            if (top == "0" && propertiesTransform.Offset.Y != null && propertiesTransform.Offset.Y.HasValue && propertiesTransform.Offset.Y.Value != 0)
                            {
                                elementStylesTransforms += $" translateY({GetInvariantNumber(RoundNumber(propertiesTransform.Offset.Y.Value / 914400.0 * 96, config.RoundingDigits))}px)";
                            }
                        }
                        if (propertiesTransform.Extents != null)
                        {
                            if (widthActual == "auto" && propertiesTransform.Extents.Cx != null && propertiesTransform.Extents.Cx.HasValue)
                            {
                                widthActual = $"{GetInvariantNumber(RoundNumber(propertiesTransform.Extents.Cx.Value / 914400.0 * 96, config.RoundingDigits))}px";
                            }
                            if (heightActual == "auto" && propertiesTransform.Extents.Cy != null && propertiesTransform.Extents.Cy.HasValue)
                            {
                                heightActual = $"{GetInvariantNumber(RoundNumber(propertiesTransform.Extents.Cy.Value / 914400.0 * 96, config.RoundingDigits))}px";
                            }
                        }
                        if (propertiesTransform.Rotation != null && propertiesTransform.Rotation.HasValue && propertiesTransform.Rotation.Value != 0)
                        {
                            elementStylesTransforms += $" rotate({GetInvariantNumber(RoundNumber(propertiesTransform.Rotation.Value / 60000.0, config.RoundingDigits))}deg)";
                        }
                        if (propertiesTransform.HorizontalFlip != null && (!propertiesTransform.HorizontalFlip.HasValue || propertiesTransform.HorizontalFlip.Value))
                        {
                            elementStylesTransforms += $" scaleX(-1)";
                        }
                        if (propertiesTransform.VerticalFlip != null && (!propertiesTransform.VerticalFlip.HasValue || propertiesTransform.VerticalFlip.Value))
                        {
                            elementStylesTransforms += $" scaleY(-1)";
                        }
                        if (!string.IsNullOrEmpty(elementStylesTransforms))
                        {
                            elementStyles["transform"] = elementStylesTransforms.TrimStart();
                        }
                    }
                    else if (propertiesElement is DocumentFormat.OpenXml.Drawing.NoFill)
                    {
                        isFillHandled = true;
                    }
                    else if (propertiesElement is DocumentFormat.OpenXml.Drawing.SolidFill propertiesFillSolid && GetColor(propertiesFillSolid, out string fillColor, themes, config))
                    {
                        elementStyles["background"] = fillColor;
                        isFillHandled = true;
                    }
                    else if (propertiesElement is DocumentFormat.OpenXml.Drawing.Outline propertiesOutline)
                    {
                        string outlineWidth = propertiesOutline.Width != null && propertiesOutline.Width.HasValue ? $"{GetInvariantNumber(RoundNumber(propertiesOutline.Width.Value / 914400.0 * 96, config.RoundingDigits))}px" : "thin";
                        string outlineStyle = propertiesOutline.CompoundLineType != null && propertiesOutline.CompoundLineType.HasValue && propertiesOutline.CompoundLineType.Value != DocumentFormat.OpenXml.Drawing.CompoundLineValues.Single ? "double" : "solid";
                        string outlineColor = string.Empty;
                        foreach (OpenXmlElement outlineElement in propertiesOutline.Elements())
                        {
                            if (outlineElement is DocumentFormat.OpenXml.Drawing.PresetDash outlineDash && outlineDash.Val != null && outlineDash.Val.HasValue && outlineDash.Val.Value != DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid)
                            {
                                outlineStyle = outlineDash.Val.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot || outlineDash.Val.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.SystemDashDotDot ? "dotted" : "dashed";
                            }
                            else if (outlineElement is DocumentFormat.OpenXml.Drawing.CustomDash)
                            {
                                outlineStyle = "dashed";
                            }
                            else if (outlineElement is DocumentFormat.OpenXml.Drawing.SolidFill outlineFillSolid)
                            {
                                GetColor(outlineFillSolid, out outlineColor, themes, config);
                            }
                        }
                        elementStyles["border"] = $"{outlineWidth} {outlineStyle}{(!string.IsNullOrEmpty(outlineColor) ? $" {outlineColor}" : string.Empty)}";
                        isOutlineHandled = true;
                    }
                    else if (propertiesElement is DocumentFormat.OpenXml.Drawing.PresetGeometry geometryPreset && geometryPreset.Preset != null && geometryPreset.Preset.HasValue && geometryPreset.Preset.Value != DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle)
                    {
                        //TODO: preset shapes
                    }
                    else if (propertiesElement is DocumentFormat.OpenXml.Drawing.CustomGeometry geometryCustom && geometryCustom.PathList != null)
                    {
                        if (geometryCustom.Rectangle != null)
                        {
                            elementPaddingTop += geometryCustom.Rectangle.Top != null && geometryCustom.Rectangle.Top.HasValue && double.TryParse(geometryCustom.Rectangle.Top.Value, out double rectangleTop) ? rectangleTop / 914400.0 * 96 : 0;
                            elementPaddingRight += geometryCustom.Rectangle.Right != null && geometryCustom.Rectangle.Right.HasValue && double.TryParse(geometryCustom.Rectangle.Right.Value, out double rectangleRight) ? rectangleRight / 914400.0 * 96 : 0;
                            elementPaddingBottom += geometryCustom.Rectangle.Bottom != null && geometryCustom.Rectangle.Bottom.HasValue && double.TryParse(geometryCustom.Rectangle.Bottom.Value, out double rectangleBottom) ? rectangleBottom / 914400.0 * 96 : 0;
                            elementPaddingLeft += geometryCustom.Rectangle.Left != null && geometryCustom.Rectangle.Left.HasValue && double.TryParse(geometryCustom.Rectangle.Left.Value, out double rectangleLeft) ? rectangleLeft / 914400.0 * 96 : 0;
                        }
                        string attribute = $"path('";
                        double pointLastX = 0;
                        double pointLastY = 0;
                        foreach (OpenXmlElement geometryPathElement in geometryCustom.PathList.Elements<DocumentFormat.OpenXml.Drawing.Path>().SelectMany(x => x.Elements()))
                        {
                            if (geometryPathElement is DocumentFormat.OpenXml.Drawing.ArcTo geometryPathArcTo)
                            {
                                double arcRadiusX = geometryPathArcTo.WidthRadius != null && geometryPathArcTo.WidthRadius.HasValue && double.TryParse(geometryPathArcTo.WidthRadius.Value, out double radiusWidth) ? radiusWidth / 914400.0 * 96 : 0;
                                double arcRadiusY = geometryPathArcTo.HeightRadius != null && geometryPathArcTo.HeightRadius.HasValue && double.TryParse(geometryPathArcTo.HeightRadius.Value, out double radiusHeight) ? radiusHeight / 914400.0 * 96 : 0;
                                double arcAngleStart = geometryPathArcTo.StartAngle != null && geometryPathArcTo.StartAngle.HasValue && double.TryParse(geometryPathArcTo.StartAngle.Value, out double angleStart) ? angleStart / 60000.0 * Math.PI / 180 : 0;
                                double arcAngleEnd = geometryPathArcTo.SwingAngle != null && geometryPathArcTo.SwingAngle.HasValue && double.TryParse(geometryPathArcTo.SwingAngle.Value, out double angleSwing) ? arcAngleStart + angleSwing / 60000.0 * Math.PI / 180 : arcAngleStart;
                                pointLastX = RoundNumber(pointLastX - arcRadiusX * Math.Cos(arcAngleStart) + arcRadiusX * Math.Cos(arcAngleEnd), config.RoundingDigits);
                                pointLastY = RoundNumber(pointLastY - arcRadiusY * Math.Sin(arcAngleStart) + arcRadiusY * Math.Sin(arcAngleEnd), config.RoundingDigits);
                                attribute += $"A {GetInvariantNumber(RoundNumber(arcRadiusX, config.RoundingDigits))} {GetInvariantNumber(RoundNumber(arcRadiusY, config.RoundingDigits))} 0 1 1 {GetInvariantNumber(pointLastX)},{GetInvariantNumber(pointLastY)} ";
                            }
                            else if (geometryPathElement is DocumentFormat.OpenXml.Drawing.CloseShapePath)
                            {
                                attribute += "Z ";
                            }
                            else
                            {
                                if (geometryPathElement is DocumentFormat.OpenXml.Drawing.MoveTo)
                                {
                                    attribute += "M ";
                                }
                                else if (geometryPathElement is DocumentFormat.OpenXml.Drawing.LineTo)
                                {
                                    attribute += "L ";
                                }
                                else if (geometryPathElement is DocumentFormat.OpenXml.Drawing.CubicBezierCurveTo)
                                {
                                    attribute += "C ";
                                }
                                else if (geometryPathElement is DocumentFormat.OpenXml.Drawing.QuadraticBezierCurveTo)
                                {
                                    attribute += "Q ";
                                }
                                foreach (DocumentFormat.OpenXml.Drawing.Point geometryPathPoint in geometryPathElement.Elements<DocumentFormat.OpenXml.Drawing.Point>())
                                {
                                    pointLastX = geometryPathPoint.X != null && geometryPathPoint.X.HasValue && double.TryParse(geometryPathPoint.X.Value, out double pointX) ? RoundNumber(pointX / 914400.0 * 96, config.RoundingDigits) : 0;
                                    pointLastY = geometryPathPoint.Y != null && geometryPathPoint.Y.HasValue && double.TryParse(geometryPathPoint.Y.Value, out double pointY) ? RoundNumber(pointY / 914400.0 * 96, config.RoundingDigits) : 0;
                                    attribute += $"{GetInvariantNumber(pointLastX)},{GetInvariantNumber(pointLastY)} ";
                                }
                            }
                        }
                        elementStyles["clip-path"] = $"{attribute.TrimEnd()}')";
                    }
                }
            }
            if (elementShapeStyle != null)
            {
                if (!isFillHandled && elementShapeStyle.FillReference != null && GetColor(elementShapeStyle.FillReference, out string fillColor, themes, config))
                {
                    elementStyles["background"] = fillColor;
                }
                if (!isOutlineHandled && elementShapeStyle.LineReference != null && GetColor(elementShapeStyle.LineReference, out string outlineColor, themes, config))
                {
                    elementStyles["border"] = $"thin solid {outlineColor}";
                }
                if (!elementStyles.ContainsKey("color") && elementShapeStyle.FontReference != null && GetColor(elementShapeStyle.FontReference, out string fontColor, themes, config))
                {
                    elementStyles["color"] = fontColor;
                }
            }
            if (elementNonVisualProperties != null && elementNonVisualProperties.Hidden != null && (!elementNonVisualProperties.Hidden.HasValue || elementNonVisualProperties.Hidden.Value))
            {
                elementStyles["visibility"] = "hidden";
            }
            if (elementPaddingTop != 0 || elementPaddingRight != 0 || elementPaddingBottom != 0 || elementPaddingLeft != 0)
            {
                elementStyles["padding"] = $"{GetInvariantNumber(RoundNumber(elementPaddingTop, config.RoundingDigits))}px {GetInvariantNumber(RoundNumber(elementPaddingRight, config.RoundingDigits))}px {GetInvariantNumber(RoundNumber(elementPaddingBottom, config.RoundingDigits))}px {GetInvariantNumber(RoundNumber(elementPaddingLeft, config.RoundingDigits))}px";
            }
            return element.Replace("{0}", $"position: absolute; left: {left}; top: {top}; width: {widthActual}; height: {heightActual};{GetAttributesString(elementStyles, true, -1)}");
        }

        private static bool GetColor(OpenXmlElement color, out string result, DocumentFormat.OpenXml.Drawing.Color2Type[] themes, ConverterConfig config)
        {
            if (color == null)
            {
                result = string.Empty;
                return false;
            }

            void actionConvertHexToRgba(string hex, out byte red, out byte green, out byte blue, out int alpha)
            {
                string hexTrimmed = hex.TrimStart('#');
                red = (byte)(hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 2 : 0, 2), 16) : 0);
                green = (byte)(hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 4 : 2, 2), 16) : 0);
                blue = (byte)(hexTrimmed.Length > 5 ? Convert.ToInt32(hexTrimmed.Substring(hexTrimmed.Length > 7 ? 6 : 4, 2), 16) : 0);
                alpha = hexTrimmed.Length > 7 ? (int)(Convert.ToInt32(hexTrimmed.Substring(0, 2), 16) / 255.0 * 100) : 100;
            }
            void actionConvertRgbToHsl(byte red, byte green, byte blue, out double hue, out double saturation, out double luminance)
            {
                double redMapped = red / 255.0;
                double greenMapped = green / 255.0;
                double blueMapped = blue / 255.0;

                double max = Math.Max(redMapped, Math.Max(greenMapped, blueMapped));
                double min = Math.Min(redMapped, Math.Min(greenMapped, blueMapped));
                double chroma = max - min;
                double distanceRed = (max - redMapped) / chroma;
                double distanceGreen = (max - greenMapped) / chroma;
                double distanceBlue = (max - blueMapped) / chroma;
                hue = max != min ? ((redMapped == max ? distanceBlue - distanceGreen : (greenMapped == max ? 2 + distanceRed - distanceBlue : 4 + distanceGreen - distanceRed)) * 60 % 360 + 360) % 360 : 0;
                saturation = max != min ? (chroma / (max + min > 1 ? 2 - max - min : max + min)) : 0;
                luminance = (max + min) / 2;
            }
            void actionConvertHslToRgb(double hue, double saturation, double luminance, out byte red, out byte green, out byte blue)
            {
                double value1 = luminance <= 0.5 ? luminance * (1 + saturation) : luminance + saturation - luminance * saturation;
                double value2 = 2 * luminance - value1;
                double actionCalculateColor(double hueShifted)
                {
                    hueShifted = (hueShifted % 360 + 360) % 360;
                    return hueShifted < 60 ? value2 + (value1 - value2) * hueShifted / 60 : (hueShifted < 180 ? value1 : (hueShifted < 240 ? value2 + (value1 - value2) * (240 - hueShifted) / 60 : value2));
                }
                red = (byte)((saturation == 0 ? luminance : actionCalculateColor(hue + 120)) * 255);
                green = (byte)((saturation == 0 ? luminance : actionCalculateColor(hue)) * 255);
                blue = (byte)((saturation == 0 ? luminance : actionCalculateColor(hue - 120)) * 255);
            }

            byte resultRed = 0;
            byte resultGreen = 0;
            byte resultBlue = 0;
            int resultAlpha = 100;
            bool actionGetColorElement(OpenXmlElement colorElement, OpenXmlElementList effects)
            {
                if (colorElement == null)
                {
                    return false;
                }

                if (colorElement is DocumentFormat.OpenXml.Drawing.RgbColorModelHex colorRgbHex && colorRgbHex.Val != null && colorRgbHex.Val.HasValue)
                {
                    actionConvertHexToRgba(colorRgbHex.Val.Value, out resultRed, out resultGreen, out resultBlue, out resultAlpha);
                }
                else if (colorElement is DocumentFormat.OpenXml.Drawing.RgbColorModelPercentage colorRgbPercentage)
                {
                    resultRed = (byte)(colorRgbPercentage.RedPortion != null && colorRgbPercentage.RedPortion.HasValue ? colorRgbPercentage.RedPortion.Value / 100000.0 * 255 : 0);
                    resultGreen = (byte)(colorRgbPercentage.GreenPortion != null && colorRgbPercentage.GreenPortion.HasValue ? colorRgbPercentage.GreenPortion.Value / 100000.0 * 255 : 0);
                    resultBlue = (byte)(colorRgbPercentage.BluePortion != null && colorRgbPercentage.BluePortion.HasValue ? colorRgbPercentage.BluePortion.Value / 100000.0 * 255 : 0);
                }
                else if (colorElement is DocumentFormat.OpenXml.Drawing.HslColor colorHsl)
                {
                    double hue = colorHsl.HueValue != null && colorHsl.HueValue.HasValue ? colorHsl.HueValue.Value / 60000.0 : 0;
                    double saturation = colorHsl.SatValue != null && colorHsl.SatValue.HasValue ? colorHsl.SatValue.Value / 100000.0 : 0;
                    double luminance = colorHsl.LumValue != null && colorHsl.LumValue.HasValue ? colorHsl.LumValue.Value / 100000.0 : 0;
                    actionConvertHslToRgb(hue, saturation, luminance, out resultRed, out resultGreen, out resultBlue);
                }
                else if (colorElement is DocumentFormat.OpenXml.Drawing.SystemColor colorSystem)
                {
                    if (colorSystem.Val != null && colorSystem.Val.HasValue && themeSystemColors.ContainsKey(colorSystem.Val.Value))
                    {
                        byte[] colorDictionary = themeSystemColors[colorSystem.Val.Value];
                        resultRed = colorDictionary[0];
                        resultGreen = colorDictionary[1];
                        resultBlue = colorDictionary[2];
                    }
                    else if (colorSystem.LastColor != null && colorSystem.LastColor.HasValue)
                    {
                        actionConvertHexToRgba(colorSystem.LastColor.Value, out resultRed, out resultGreen, out resultBlue, out resultAlpha);
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (colorElement is DocumentFormat.OpenXml.Drawing.PresetColor colorPreset && colorPreset.Val != null && colorPreset.Val.HasValue && themePresetColors.ContainsKey(colorPreset.Val.Value))
                {
                    byte[] colorDictionary = themePresetColors[colorPreset.Val.Value];
                    resultRed = colorDictionary[0];
                    resultGreen = colorDictionary[1];
                    resultBlue = colorDictionary[2];
                }
                else if (colorElement is DocumentFormat.OpenXml.Drawing.SchemeColor colorScheme && colorScheme.Val != null && colorScheme.Val.HasValue)
                {
                    DocumentFormat.OpenXml.Drawing.Color2Type colorTheme = null;
                    if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Light1)
                    {
                        colorTheme = themes[0];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Dark1)
                    {
                        colorTheme = themes[1];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Light2)
                    {
                        colorTheme = themes[2];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Dark2)
                    {
                        colorTheme = themes[3];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1)
                    {
                        colorTheme = themes[4];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent2)
                    {
                        colorTheme = themes[5];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent3)
                    {
                        colorTheme = themes[6];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent4)
                    {
                        colorTheme = themes[7];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent5)
                    {
                        colorTheme = themes[8];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent6)
                    {
                        colorTheme = themes[9];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Hyperlink)
                    {
                        colorTheme = themes[10];
                    }
                    else if (colorScheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.FollowedHyperlink)
                    {
                        colorTheme = themes[11];
                    }
                    return colorTheme != null && colorTheme.FirstChild is OpenXmlElement colorThemeChildElement && actionGetColorElement(colorThemeChildElement, colorScheme.ChildElements);
                }
                else
                {
                    return false;
                }
                foreach (OpenXmlElement effect in effects)
                {
                    if (effect is DocumentFormat.OpenXml.Drawing.Shade shade && shade.Val != null && shade.Val.HasValue)
                    {
                        double amount = shade.Val.Value / 100000.0;
                        resultRed = (byte)(resultRed * amount);
                        resultGreen = (byte)(resultGreen * amount);
                        resultBlue = (byte)(resultBlue * amount);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Tint tint && tint.Val != null && tint.Val.HasValue)
                    {
                        double amount = tint.Val.Value / 100000.0;
                        resultRed = (byte)(resultRed * amount + 255 * (1 - amount));
                        resultGreen = (byte)(resultGreen * amount + 255 * (1 - amount));
                        resultBlue = (byte)(resultBlue * amount + 255 * (1 - amount));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Inverse)
                    {
                        resultRed = (byte)(255 - resultRed);
                        resultGreen = (byte)(255 - resultGreen);
                        resultBlue = (byte)(255 - resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Gray)
                    {
                        byte grayscale = (byte)(resultRed * 0.3 + resultGreen * 0.59 + resultBlue * 0.11);
                        resultRed = grayscale;
                        resultGreen = grayscale;
                        resultBlue = grayscale;
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Complement)
                    {
                        double colorMax = Math.Max(resultRed, Math.Max(resultGreen, resultBlue));
                        resultRed = (byte)(colorMax - resultRed);
                        resultGreen = (byte)(colorMax - resultGreen);
                        resultBlue = (byte)(colorMax - resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Gamma)
                    {
                        double resultRedMapped = resultRed / 255.0;
                        resultRed = (byte)((resultRedMapped > 0.04045 ? Math.Pow((resultRedMapped + 0.055) / 1.055, 2.4) : resultRedMapped / 12.92) * 255);
                        double resultGreenMapped = resultGreen / 255.0;
                        resultGreen = (byte)((resultGreenMapped > 0.04045 ? Math.Pow((resultGreenMapped + 0.055) / 1.055, 2.4) : resultGreenMapped / 12.92) * 255);
                        double resultBlueMapped = resultBlue / 255.0;
                        resultBlue = (byte)((resultBlueMapped > 0.04045 ? Math.Pow((resultBlueMapped + 0.055) / 1.055, 2.4) : resultBlueMapped / 12.92) * 255);

                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.InverseGamma)
                    {
                        double resultRedMapped = resultRed / 255.0;
                        resultRed = (byte)((resultRedMapped > 0.0031308 ? 1.055 * Math.Pow(resultRedMapped, 1 / 2.4) - 0.055 : resultRedMapped * 12.92) * 255);
                        double resultGreenMapped = resultGreen / 255.0;
                        resultGreen = (byte)((resultGreenMapped > 0.0031308 ? 1.055 * Math.Pow(resultGreenMapped, 1 / 2.4) - 0.055 : resultGreenMapped * 12.92) * 255);
                        double resultBlueMapped = resultBlue / 255.0;
                        resultBlue = (byte)((resultBlueMapped > 0.0031308 ? 1.055 * Math.Pow(resultBlueMapped, 1 / 2.4) - 0.055 : resultBlueMapped * 12.92) * 255);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Red red && red.Val != null && red.Val.HasValue)
                    {
                        resultRed = (byte)(red.Val.Value / 100000.0 * 255);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.RedModulation redModulation && redModulation.Val != null && redModulation.Val.HasValue)
                    {
                        resultRed = (byte)(resultRed * (redModulation.Val.Value / 100000.0));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.RedOffset redOffset && redOffset.Val != null && redOffset.Val.HasValue)
                    {
                        resultRed = (byte)(resultRed + (redOffset.Val.Value / 100000.0 * 255));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Green green && green.Val != null && green.Val.HasValue)
                    {
                        resultGreen = (byte)(green.Val.Value / 100000.0 * 255);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.GreenModulation greenModulation && greenModulation.Val != null && greenModulation.Val.HasValue)
                    {
                        resultGreen = (byte)(resultGreen * (greenModulation.Val.Value / 100000.0));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.GreenOffset greenOffset && greenOffset.Val != null && greenOffset.Val.HasValue)
                    {
                        resultGreen = (byte)(resultGreen + (greenOffset.Val.Value / 100000.0 * 255));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Blue blue && blue.Val != null && blue.Val.HasValue)
                    {
                        resultBlue = (byte)(blue.Val.Value / 100000.0 * 255);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.BlueModulation blueModulation && blueModulation.Val != null && blueModulation.Val.HasValue)
                    {
                        resultBlue = (byte)(resultBlue * (blueModulation.Val.Value / 100000.0));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.BlueOffset blueOffset && blueOffset.Val != null && blueOffset.Val.HasValue)
                    {
                        resultBlue = (byte)(resultBlue + (blueOffset.Val.Value / 100000.0 * 255));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Alpha alpha && alpha.Val != null && alpha.Val.HasValue)
                    {
                        resultAlpha = (int)(alpha.Val.Value / 100000.0 * 100);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.AlphaModulation alphaModulation && alphaModulation.Val != null && alphaModulation.Val.HasValue)
                    {
                        resultAlpha = (int)(resultAlpha * (alphaModulation.Val.Value / 100000.0));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.AlphaOffset alphaOffset && alphaOffset.Val != null && alphaOffset.Val.HasValue)
                    {
                        resultAlpha = (int)(resultAlpha + (alphaOffset.Val.Value / 100000.0 * 100));
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Hue hue && hue.Val != null && hue.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double _, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(hue.Val.Value / 60000.0, colorSaturation, colorLuminance, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.HueModulation hueModulation && hueModulation.Val != null && hueModulation.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(colorHue * (hueModulation.Val.Value / 100000.0), colorSaturation, colorLuminance, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.HueOffset hueOffset && hueOffset.Val != null && hueOffset.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(colorHue + hueOffset.Val.Value / 60000.0, colorSaturation, colorLuminance, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Saturation saturation && saturation.Val != null && saturation.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double _, out double colorLuminance);
                        actionConvertHslToRgb(colorHue, saturation.Val.Value / 100000.0, colorLuminance, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.SaturationModulation saturationModulation && saturationModulation.Val != null && saturationModulation.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(colorHue, colorSaturation * (saturationModulation.Val.Value / 100000.0), colorLuminance, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.SaturationOffset saturationOffset && saturationOffset.Val != null && saturationOffset.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(colorHue, colorSaturation + saturationOffset.Val.Value / 100000.0, colorLuminance, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.Luminance luminance && luminance.Val != null && luminance.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturaion, out double _);
                        actionConvertHslToRgb(colorHue, colorSaturaion, luminance.Val.Value / 100000.0, out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.LuminanceModulation luminanceModulation && luminanceModulation.Val != null && luminanceModulation.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(colorHue, colorSaturation, colorLuminance * (luminanceModulation.Val.Value / 100000.0), out resultRed, out resultGreen, out resultBlue);
                    }
                    else if (effect is DocumentFormat.OpenXml.Drawing.LuminanceOffset luminanceOffset && luminanceOffset.Val != null && luminanceOffset.Val.HasValue)
                    {
                        actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double colorHue, out double colorSaturation, out double colorLuminance);
                        actionConvertHslToRgb(colorHue, colorSaturation, colorLuminance + luminanceOffset.Val.Value / 100000.0, out resultRed, out resultGreen, out resultBlue);
                    }
                }
                return true;
            }

            if (color is ColorType colorType)
            {
                if (colorType.Auto != null && (!colorType.Auto.HasValue || colorType.Auto.Value))
                {
                    result = "initial";
                    return true;
                }
                else if (colorType.Rgb != null && colorType.Rgb.HasValue)
                {
                    actionConvertHexToRgba(colorType.Rgb.Value, out resultRed, out resultGreen, out resultBlue, out resultAlpha);
                }
                else if (colorType.Indexed != null && colorType.Indexed.HasValue)
                {
                    void actionUpdateColor(byte red, byte green, byte blue)
                    {
                        resultRed = red;
                        resultGreen = green;
                        resultBlue = blue;
                    };
                    switch (colorType.Indexed.Value)
                    {
                        case 0:
                            actionUpdateColor(0, 0, 0);
                            break;
                        case 1:
                            actionUpdateColor(255, 255, 255);
                            break;
                        case 2:
                            actionUpdateColor(255, 0, 0);
                            break;
                        case 3:
                            actionUpdateColor(0, 255, 0);
                            break;
                        case 4:
                            actionUpdateColor(0, 0, 255);
                            break;
                        case 5:
                            actionUpdateColor(255, 255, 0);
                            break;
                        case 6:
                            actionUpdateColor(255, 0, 255);
                            break;
                        case 7:
                            actionUpdateColor(0, 255, 255);
                            break;
                        case 8:
                            actionUpdateColor(0, 0, 0);
                            break;
                        case 9:
                            actionUpdateColor(255, 255, 255);
                            break;
                        case 10:
                            actionUpdateColor(255, 0, 0);
                            break;
                        case 11:
                            actionUpdateColor(0, 255, 0);
                            break;
                        case 12:
                            actionUpdateColor(0, 0, 255);
                            break;
                        case 13:
                            actionUpdateColor(255, 255, 0);
                            break;
                        case 14:
                            actionUpdateColor(255, 0, 255);
                            break;
                        case 15:
                            actionUpdateColor(0, 255, 255);
                            break;
                        case 16:
                            actionUpdateColor(128, 0, 0);
                            break;
                        case 17:
                            actionUpdateColor(0, 128, 0);
                            break;
                        case 18:
                            actionUpdateColor(0, 0, 128);
                            break;
                        case 19:
                            actionUpdateColor(128, 128, 0);
                            break;
                        case 20:
                            actionUpdateColor(128, 0, 128);
                            break;
                        case 21:
                            actionUpdateColor(0, 128, 128);
                            break;
                        case 22:
                            actionUpdateColor(192, 192, 192);
                            break;
                        case 23:
                            actionUpdateColor(128, 128, 128);
                            break;
                        case 24:
                            actionUpdateColor(153, 153, 255);
                            break;
                        case 25:
                            actionUpdateColor(153, 51, 102);
                            break;
                        case 26:
                            actionUpdateColor(255, 255, 204);
                            break;
                        case 27:
                            actionUpdateColor(204, 255, 255);
                            break;
                        case 28:
                            actionUpdateColor(102, 0, 102);
                            break;
                        case 29:
                            actionUpdateColor(255, 128, 128);
                            break;
                        case 30:
                            actionUpdateColor(0, 102, 204);
                            break;
                        case 31:
                            actionUpdateColor(204, 204, 255);
                            break;
                        case 32:
                            actionUpdateColor(0, 0, 128);
                            break;
                        case 33:
                            actionUpdateColor(255, 0, 255);
                            break;
                        case 34:
                            actionUpdateColor(255, 255, 0);
                            break;
                        case 35:
                            actionUpdateColor(0, 255, 255);
                            break;
                        case 36:
                            actionUpdateColor(128, 0, 128);
                            break;
                        case 37:
                            actionUpdateColor(128, 0, 0);
                            break;
                        case 38:
                            actionUpdateColor(0, 128, 128);
                            break;
                        case 39:
                            actionUpdateColor(0, 0, 255);
                            break;
                        case 40:
                            actionUpdateColor(0, 204, 255);
                            break;
                        case 41:
                            actionUpdateColor(204, 255, 255);
                            break;
                        case 42:
                            actionUpdateColor(204, 255, 204);
                            break;
                        case 43:
                            actionUpdateColor(255, 255, 153);
                            break;
                        case 44:
                            actionUpdateColor(153, 204, 255);
                            break;
                        case 45:
                            actionUpdateColor(255, 153, 204);
                            break;
                        case 46:
                            actionUpdateColor(204, 153, 255);
                            break;
                        case 47:
                            actionUpdateColor(255, 204, 153);
                            break;
                        case 48:
                            actionUpdateColor(51, 102, 255);
                            break;
                        case 49:
                            actionUpdateColor(51, 204, 204);
                            break;
                        case 50:
                            actionUpdateColor(153, 204, 0);
                            break;
                        case 51:
                            actionUpdateColor(255, 204, 0);
                            break;
                        case 52:
                            actionUpdateColor(255, 153, 0);
                            break;
                        case 53:
                            actionUpdateColor(255, 102, 0);
                            break;
                        case 54:
                            actionUpdateColor(102, 102, 153);
                            break;
                        case 55:
                            actionUpdateColor(150, 150, 150);
                            break;
                        case 56:
                            actionUpdateColor(0, 51, 102);
                            break;
                        case 57:
                            actionUpdateColor(51, 153, 102);
                            break;
                        case 58:
                            actionUpdateColor(0, 51, 0);
                            break;
                        case 59:
                            actionUpdateColor(51, 51, 0);
                            break;
                        case 60:
                            actionUpdateColor(153, 51, 0);
                            break;
                        case 61:
                            actionUpdateColor(153, 51, 102);
                            break;
                        case 62:
                            actionUpdateColor(51, 51, 153);
                            break;
                        case 63:
                            actionUpdateColor(51, 51, 51);
                            break;
                        case 64:
                            actionUpdateColor(128, 128, 128);
                            break;
                        case 65:
                            actionUpdateColor(255, 255, 255);
                            break;
                        default:
                            result = string.Empty;
                            return false;
                    }
                }
                else if (colorType.Theme == null || !colorType.Theme.HasValue || themes == null || colorType.Theme.Value < 0 || colorType.Theme.Value >= themes.Length || !(themes[colorType.Theme.Value].FirstChild is OpenXmlElement themeColorElement) || !actionGetColorElement(themeColorElement, themeColorElement.ChildElements))
                {
                    result = string.Empty;
                    return false;
                }
                if (colorType.Tint != null && colorType.Tint.HasValue && colorType.Tint.Value != 0)
                {
                    actionConvertRgbToHsl(resultRed, resultGreen, resultBlue, out double hue, out double saturation, out double luminance);
                    actionConvertHslToRgb(hue, saturation, colorType.Tint.Value < 0 ? luminance * (1 + colorType.Tint.Value) : luminance * (1 - colorType.Tint.Value) + colorType.Tint.Value, out resultRed, out resultGreen, out resultBlue);
                }
            }
            else if (!(color.FirstChild is OpenXmlElement colorElement) || !actionGetColorElement(colorElement, colorElement.ChildElements))
            {
                result = string.Empty;
                return false;
            }
            result = config.UseHexColors ? $"#{resultRed:X2}{resultGreen:X2}{resultBlue:X2}{(resultAlpha < 100 ? ((byte)(resultAlpha / 100.0 * 255)).ToString("X2") : string.Empty)}" : $"{(resultAlpha < 100 ? "rgba" : "rgb")}({resultRed}, {resultGreen}, {resultBlue}{(resultAlpha < 100 ? $", {Math.Round(resultAlpha / 100.0, 2)}" : string.Empty)})";
            return true;
        }

        #endregion

        #region Private Fields

        private const string stringFormatNumber = "0.##############################";

        private static readonly Regex regexNumbers = new Regex(@"\d+", RegexOptions.Compiled);
        private static readonly Regex regexLetters = new Regex("[A-Za-z]+", RegexOptions.Compiled);

        private static readonly Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, byte[]> themeSystemColors = new Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, byte[]>()
        {
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveBorder, new byte[3] { 180, 180, 180 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveCaption, new byte[3] { 153, 180, 209 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ApplicationWorkspace, new byte[3] { 171, 171, 171 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Background, new byte[3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonFace, new byte[3] { 240, 240, 240 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonHighlight, new byte[3] { 0, 120, 215 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonShadow, new byte[3] { 160, 160, 160 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonText, new byte[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.CaptionText, new byte[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientActiveCaption, new byte[3] { 185, 209, 234 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientInactiveCaption, new byte[3] { 215, 228, 242 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.GrayText, new byte[3] { 109, 109, 109 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Highlight, new byte[3] { 0, 120, 215 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.HighlightText, new byte[3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.HotLight, new byte[3] { 255, 165, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveBorder, new byte[3] { 244, 247, 252 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaption, new byte[3] { 191, 205, 219 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaptionText, new byte[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoBack, new byte[3] { 255, 255, 225 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoText, new byte[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Menu, new byte[3] { 240, 240, 240 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuBar, new byte[3] { 240, 240, 240 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuHighlight, new byte[3] { 0, 120, 215 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuText, new byte[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ScrollBar, new byte[3] { 200, 200, 200 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDDarkShadow, new byte[3] { 160, 160, 160 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDLight, new byte[3] { 227, 227, 227 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.Window, new byte[3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowFrame, new byte[3] { 100, 100, 100 } },
            { DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText, new byte[3] { 0, 0, 0 } }
        };

        private static readonly Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, byte[]> themePresetColors = new Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, byte[]>()
        {
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.AliceBlue, new byte[3] { 240, 248, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.AntiqueWhite, new byte[3] { 250, 235, 215 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Aqua, new byte[3] { 0, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Aquamarine, new byte[3] { 127, 255, 212 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Azure, new byte[3] { 240, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Beige, new byte[3] { 245, 245, 220 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Bisque, new byte[3] { 255, 228, 196 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Black, new byte[3] { 0, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.BlanchedAlmond, new byte[3] { 255, 235, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Blue, new byte[3] { 0, 0, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.BlueViolet, new byte[3] { 138, 43, 226 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Brown, new byte[3] { 165, 42, 42 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.BurlyWood, new byte[3] { 222, 184, 135 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.CadetBlue, new byte[3] { 95, 158, 160 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Chartreuse, new byte[3] { 127, 255, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Chocolate, new byte[3] { 210, 105, 30 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Coral, new byte[3] { 255, 127, 80 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.CornflowerBlue, new byte[3] { 100, 149, 237 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Cornsilk, new byte[3] { 255, 248, 220 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Crimson, new byte[3] { 220, 20, 60 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Cyan, new byte[3] { 0, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue, new byte[3] { 0, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan, new byte[3] { 0, 139, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod, new byte[3] { 184, 134, 11 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray, new byte[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen, new byte[3] { 0, 100, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki, new byte[3] { 189, 183, 107 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta, new byte[3] { 139, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen, new byte[3] { 85, 107, 47 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange, new byte[3] { 255, 140, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid, new byte[3] { 153, 50, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed, new byte[3] { 139, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon, new byte[3] { 233, 150, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen, new byte[3] { 143, 188, 143 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue, new byte[3] { 72, 61, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray, new byte[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise, new byte[3] { 0, 206, 209 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet, new byte[3] { 148, 0, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepPink, new byte[3] { 255, 20, 147 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepSkyBlue, new byte[3] { 0, 191, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGray, new byte[3] { 105, 105, 105 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DodgerBlue, new byte[3] { 30, 144, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Firebrick, new byte[3] { 178, 34, 34 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.FloralWhite, new byte[3] { 255, 250, 240 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.ForestGreen, new byte[3] { 34, 139, 34 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Fuchsia, new byte[3] { 255, 0, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Gainsboro, new byte[3] { 220, 220, 220 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.GhostWhite, new byte[3] { 248, 248, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Gold, new byte[3] { 255, 215, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Goldenrod, new byte[3] { 218, 165, 32 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Gray, new byte[3] { 128, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Green, new byte[3] { 0, 128, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.GreenYellow, new byte[3] { 173, 255, 47 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Honeydew, new byte[3] { 240, 255, 240 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.HotPink, new byte[3] { 255, 105, 180 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.IndianRed, new byte[3] { 205, 92, 92 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Indigo, new byte[3] { 75, 0, 130 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Ivory, new byte[3] { 255, 255, 240 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Khaki, new byte[3] { 240, 230, 140 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Lavender, new byte[3] { 230, 230, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LavenderBlush, new byte[3] { 255, 240, 245 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LawnGreen, new byte[3] { 124, 252, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LemonChiffon, new byte[3] { 255, 250, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue, new byte[3] { 173, 216, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral, new byte[3] { 240, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan, new byte[3] { 224, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow, new byte[3] { 250, 250, 210 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray, new byte[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen, new byte[3] { 144, 238, 144 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink, new byte[3] { 255, 182, 193 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon, new byte[3] { 255, 160, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen, new byte[3] { 32, 178, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue, new byte[3] { 135, 206, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray, new byte[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue, new byte[3] { 176, 196, 222 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow, new byte[3] { 255, 255, 224 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Lime, new byte[3] { 0, 255, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LimeGreen, new byte[3] { 50, 205, 50 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Linen, new byte[3] { 250, 240, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Magenta, new byte[3] { 255, 0, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Maroon, new byte[3] { 128, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MedAquamarine, new byte[3] { 102, 205, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue, new byte[3] { 0, 0, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid, new byte[3] { 186, 85, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple, new byte[3] { 147, 112, 219 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen, new byte[3] { 60, 179, 113 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue, new byte[3] { 123, 104, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen, new byte[3] { 0, 250, 154 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise, new byte[3] { 72, 209, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed, new byte[3] { 199, 21, 133 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MidnightBlue, new byte[3] { 25, 25, 112 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MintCream, new byte[3] { 245, 255, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MistyRose, new byte[3] { 255, 228, 225 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Moccasin, new byte[3] { 255, 228, 181 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.NavajoWhite, new byte[3] { 255, 222, 173 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Navy, new byte[3] { 0, 0, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.OldLace, new byte[3] { 253, 245, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Olive, new byte[3] { 128, 128, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.OliveDrab, new byte[3] { 107, 142, 35 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Orange, new byte[3] { 255, 165, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.OrangeRed, new byte[3] { 255, 69, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Orchid, new byte[3] { 218, 112, 214 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGoldenrod, new byte[3] { 238, 232, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGreen, new byte[3] { 152, 251, 152 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleTurquoise, new byte[3] { 175, 238, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleVioletRed, new byte[3] { 219, 112, 147 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PapayaWhip, new byte[3] { 255, 239, 213 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PeachPuff, new byte[3] { 255, 218, 185 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Peru, new byte[3] { 205, 133, 63 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Pink, new byte[3] { 255, 192, 203 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Plum, new byte[3] { 221, 160, 221 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.PowderBlue, new byte[3] { 176, 224, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Purple, new byte[3] { 128, 0, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Red, new byte[3] { 255, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.RosyBrown, new byte[3] { 188, 143, 143 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.RoyalBlue, new byte[3] { 65, 105, 225 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SaddleBrown, new byte[3] { 139, 69, 19 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Salmon, new byte[3] { 250, 128, 114 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SandyBrown, new byte[3] { 244, 164, 96 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaGreen, new byte[3] { 46, 139, 87 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaShell, new byte[3] { 255, 245, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Sienna, new byte[3] { 160, 82, 45 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Silver, new byte[3] { 192, 192, 192 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SkyBlue, new byte[3] { 135, 206, 235 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateBlue, new byte[3] { 106, 90, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGray, new byte[3] { 112, 128, 144 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Snow, new byte[3] { 255, 250, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SpringGreen, new byte[3] { 0, 255, 127 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SteelBlue, new byte[3] { 70, 130, 180 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Tan, new byte[3] { 210, 180, 140 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Teal, new byte[3] { 0, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Thistle, new byte[3] { 216, 191, 216 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Tomato, new byte[3] { 255, 99, 71 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Turquoise, new byte[3] { 64, 224, 208 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Violet, new byte[3] { 238, 130, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Wheat, new byte[3] { 245, 222, 179 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.White, new byte[3] { 255, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.WhiteSmoke, new byte[3] { 245, 245, 245 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Yellow, new byte[3] { 255, 255, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.YellowGreen, new byte[3] { 154, 205, 50 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue2010, new byte[3] { 0, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan2010, new byte[3] { 0, 139, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod2010, new byte[3] { 184, 134, 11 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray2010, new byte[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey2010, new byte[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen2010, new byte[3] { 0, 100, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki2010, new byte[3] { 189, 183, 107 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta2010, new byte[3] { 139, 0, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen2010, new byte[3] { 85, 107, 47 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange2010, new byte[3] { 255, 140, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid2010, new byte[3] { 153, 50, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed2010, new byte[3] { 139, 0, 0 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon2010, new byte[3] { 233, 150, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen2010, new byte[3] { 143, 188, 143 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue2010, new byte[3] { 72, 61, 139 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray2010, new byte[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey2010, new byte[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise2010, new byte[3] { 0, 206, 209 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet2010, new byte[3] { 148, 0, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue2010, new byte[3] { 173, 216, 230 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral2010, new byte[3] { 240, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan2010, new byte[3] { 224, 255, 255 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow2010, new byte[3] { 250, 250, 210 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray2010, new byte[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey2010, new byte[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen2010, new byte[3] { 144, 238, 144 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink2010, new byte[3] { 255, 182, 193 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon2010, new byte[3] { 255, 160, 122 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen2010, new byte[3] { 32, 178, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue2010, new byte[3] { 135, 206, 250 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray2010, new byte[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey2010, new byte[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue2010, new byte[3] { 176, 196, 222 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow2010, new byte[3] { 255, 255, 224 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumAquamarine2010, new byte[3] { 102, 205, 170 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue2010, new byte[3] { 0, 0, 205 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid2010, new byte[3] { 186, 85, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple2010, new byte[3] { 147, 112, 219 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen2010, new byte[3] { 60, 179, 113 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue2010, new byte[3] { 123, 104, 238 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen2010, new byte[3] { 0, 250, 154 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise2010, new byte[3] { 72, 209, 204 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed2010, new byte[3] { 199, 21, 133 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey, new byte[3] { 169, 169, 169 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGrey, new byte[3] { 105, 105, 105 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey, new byte[3] { 47, 79, 79 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.Grey, new byte[3] { 128, 128, 128 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey, new byte[3] { 211, 211, 211 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey, new byte[3] { 119, 136, 153 } },
            { DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGrey, new byte[3] { 112, 128, 144 } }
        };

        #endregion
    }

    /// <summary>
    /// The configurations of the Xlsx to Html converter.
    /// </summary>
    public class ConverterConfig
    {
        public const string DefaultErrorMessage = "Error! Unable to convert XLSX file! The file is either already open in another program (please close it first) or contains corrupted data.";
        public const string DefaultPresetStyles = @"
body {
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
    background: transparent;
    border: thin solid lightgray;
    border-collapse: collapse;
    white-space: nowrap;
    overflow: hidden;
    box-sizing: border-box;
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
            this.ConvertShapes = true;
            this.ConvertSheetTitles = true;
            this.ConvertHiddenSheets = false;
            this.ConvertFirstSheetOnly = false;
            this.ConvertHtmlBodyOnly = false;
            this.UseHtmlStyleClasses = true;
            this.UseHexColors = true;
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
        /// Gets or sets whether to convert Xlsx shapes to Html elements.
        /// </summary>
        public bool ConvertShapes { get; set; }

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
        /// Gets or sets whether to use the hexadecimal representations of colors.
        /// </summary>
        public bool UseHexColors { get; set; }

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
                ConvertShapes = this.ConvertShapes,
                ConvertSheetTitles = this.ConvertSheetTitles,
                ConvertHiddenSheets = this.ConvertHiddenSheets,
                ConvertFirstSheetOnly = this.ConvertFirstSheetOnly,
                ConvertHtmlBodyOnly = this.ConvertHtmlBodyOnly,
                UseHtmlStyleClasses = this.UseHtmlStyleClasses,
                UseHexColors = this.UseHexColors,
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
