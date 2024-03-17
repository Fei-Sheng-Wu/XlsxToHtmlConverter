using System;
using System.IO;
using System.Linq;
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

        #region Private Fields

        private static readonly string[] IndexedColorData = new string[] { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333", "808080", "FFFFFF" };

        #endregion

        #region Public Methods

        /// <summary>
        /// Convert a local Xlsx file to Html string.
        /// </summary>
        /// <param name="fileName">The full path to the local Xlsx file.</param>
        /// <param name="outputHtml">The output stream of the Html file.</param>
        /// <param name="loadIntoMemory">Whether to read the Xlsx file into <see cref="MemoryStream"/> at once to increase speed or not.</param>
        public static void ConvertXlsx(string fileName, Stream outputHtml, bool loadIntoMemory = true)
        {
            if (loadIntoMemory == true)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    byte[] byteArray = File.ReadAllBytes(fileName);
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    ConvertXlsx(memoryStream, outputHtml, ConverterConfig.DefaultSettings, null);
                }
            }
            else
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
                {
                    ConvertXlsx(fileStream, outputHtml, ConverterConfig.DefaultSettings, null);
                }
            }
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
            if (loadIntoMemory == true)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    byte[] byteArray = File.ReadAllBytes(fileName);
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    ConvertXlsx(memoryStream, outputHtml, config, null);
                }
            }
            else
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
                {
                    ConvertXlsx(fileStream, outputHtml, config, null);
                }
            }
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
            if (loadIntoMemory == true)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    byte[] byteArray = File.ReadAllBytes(fileName);
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    ConvertXlsx(memoryStream, outputHtml, ConverterConfig.DefaultSettings, progressCallback);
                }
            }
            else
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
                {
                    ConvertXlsx(fileStream, outputHtml, ConverterConfig.DefaultSettings, progressCallback);
                }
            }
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
            if (loadIntoMemory == true)
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
            config = config != null ? config : ConverterConfig.DefaultSettings;

            outputHtml.Seek(0, SeekOrigin.Begin);
            outputHtml.SetLength(0);

            using (StreamWriter writer = new StreamWriter(outputHtml, config.Encoding, 65536))
            {
                try
                {
                    writer.AutoFlush = true;

                    if (config.ConvertHtmlBodyOnly == false)
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
                        WorkbookStylesPart styles = workbook.WorkbookStylesPart;
                        IEnumerable<Sheet> sheets = workbook.Workbook.Descendants<Sheet>();
                        SharedStringTable sharedStringTable = workbook.GetPartsOfType<SharedStringTablePart>()?.FirstOrDefault().SharedStringTable;

                        int totalSheets = sheets.Count();
                        int progressSheetIndex = 0;
                        foreach (Sheet currentSheet in sheets)
                        {
                            progressSheetIndex++;

                            if (config.ConvertFirstSheetOnly == true && progressSheetIndex > 1)
                            {
                                continue;
                            }
                            if (config.ConvertHiddenSheet == false && currentSheet.State != null && currentSheet.State.Value != SheetStateValues.Visible)
                            {
                                continue;
                            }

                            WorksheetPart worksheet = (WorksheetPart)workbook.GetPartById(currentSheet.Id);
                            if (worksheet == null)
                            {
                                continue;
                            }
                            Worksheet sheet = worksheet.Worksheet;

                            if (config.ConvertSheetNameTitle == true)
                            {
                                writer.Write($"\n{new string(' ', 4)}<h5 {(sheet.SheetProperties != null && sheet.SheetProperties.TabColor != null ? "style=\"border-bottom-color: " + GetColorFromColorType(document, sheet.SheetProperties.TabColor, new RgbaColor() { R = 255, G = 255, B = 255, A = 0 }) + ";\"" : "")}>{currentSheet.Name}</h5>");
                            }

                            writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                            writer.Write($"\n{new string(' ', 8)}<table>");

                            double defaultRowHeight = sheet.SheetFormatProperties != null && sheet.SheetFormatProperties.DefaultRowHeight != null ? sheet.SheetFormatProperties.DefaultRowHeight.Value / 0.75 : 20;

                            bool containsMergeCells = false;
                            List<MergeCellInfo> mergeCells = new List<MergeCellInfo>();
                            if (sheet.Descendants<MergeCells>().FirstOrDefault() is MergeCells mergeCellsGroup)
                            {
                                containsMergeCells = true;

                                foreach (MergeCell mergeCell in mergeCellsGroup.Cast<MergeCell>())
                                {
                                    try
                                    {
                                        string[] range = mergeCell.Reference.Value.Split(':');

                                        uint firstColumn = GetColumnIndex(GetColumnName(range[0]));
                                        uint secondColumn = GetColumnIndex(GetColumnName(range[1]));
                                        uint firstRow = GetRowIndex(range[0]);
                                        uint secondRow = GetRowIndex(range[1]);

                                        uint fromColumn = Math.Min(firstColumn, secondColumn);
                                        uint toColumn = Math.Max(firstColumn, secondColumn);
                                        uint fromRow = Math.Min(firstRow, secondRow);
                                        uint toRow = Math.Max(firstRow, secondRow);

                                        mergeCells.Add(new MergeCellInfo() { FromColumn = fromColumn, FromRow = fromRow, ToColumn = toColumn, ToRow = toRow, ColumnSpanned = toColumn - fromColumn + 1, RowSpanned = toRow - fromRow + 1 });
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }

                            ConditionalFormatting conditionalFormatting = null;
                            if (sheet.Elements<ConditionalFormatting>() is IEnumerable<ConditionalFormatting> conditionalFormattings && conditionalFormattings.Any())
                            {
                                try
                                {
                                    conditionalFormatting = conditionalFormattings.First();
                                }
                                catch
                                {
                                    conditionalFormatting = null;
                                }
                            }

                            IEnumerable<Row> rows = sheet.Descendants<Row>();

                            uint totalRows = 0;
                            uint totalColumn = 0;
                            try
                            {
                                if (sheet.SheetDimension != null && sheet.SheetDimension.Reference != null && sheet.SheetDimension.Reference.HasValue)
                                {
                                    string[] dimension = sheet.SheetDimension.Reference.Value.Split(':');
                                    uint fromColumn = GetColumnIndex(GetColumnName(dimension[0]));
                                    uint toColumn = GetColumnIndex(GetColumnName(dimension[1]));
                                    uint fromRow = GetRowIndex(dimension[0]);
                                    uint toRow = GetRowIndex(dimension[1]);

                                    totalRows = toRow - fromColumn + 1;
                                    totalColumn = toColumn - fromColumn + 1;
                                }
                                else
                                {
                                    throw new Exception("Cannot get the sheet dimension.");
                                }
                            }
                            catch
                            {
                                foreach (Row row in rows)
                                {
                                    foreach (Cell cell in row.Descendants<Cell>())
                                    {
                                        try
                                        {
                                            string columnName = GetColumnName(cell.CellReference);
                                            uint rowIndex = GetRowIndex(cell.CellReference);

                                            uint columnIndex = GetColumnIndex(columnName.ToLower()) + 1;

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
                            }

                            int currentColumn = 0;
                            uint lastRow = 0;

                            List<double> columnWidths = new List<double>();
                            List<double> rowHeights = new List<double>();

                            for (uint i = 0; i < totalColumn; i++)
                            {
                                columnWidths.Add(double.NaN);
                            }
                            for (uint i = 0; i < totalRows; i++)
                            {
                                rowHeights.Add(defaultRowHeight);
                            }

                            if (sheet.GetFirstChild<Columns>() is Columns columnsGroup)
                            {
                                foreach (Column column in columnsGroup.Descendants<Column>())
                                {
                                    for (uint i = column.Min; i <= column.Max; i++)
                                    {
                                        try
                                        {
                                            if (column.CustomWidth && column.Width != null)
                                            {
                                                columnWidths[(int)i - 1] = (column.Width.Value - 1) * 7 + 7;
                                            }
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                    }
                                }
                            }

                            int progressRowIndex = 0;
                            foreach (Row row in rows)
                            {
                                progressRowIndex++;

                                if (row.RowIndex.Value - lastRow > 1)
                                {
                                    for (int i = 0; i < row.RowIndex.Value - lastRow - 1; i++)
                                    {
                                        writer.Write($"\n{new string(' ', 12)}<tr>");

                                        for (int j = 0; j < totalColumn; j++)
                                        {
                                            double actualCellWidth = j >= columnWidths.Count ? double.NaN : columnWidths[j];
                                            writer.Write($"\n{new string(' ', 16)}<td style=\"height: {defaultRowHeight}px; width: {(double.IsNaN(actualCellWidth) ? "auto" : actualCellWidth + "px")};\"></td>");
                                        }

                                        writer.Write($"\n{new string(' ', 12)}</tr>");
                                    }
                                }

                                currentColumn = 0;
                                double rowHeight = row.CustomHeight != null && row.CustomHeight.Value && row.Height != null ? row.Height.Value / 0.75 : defaultRowHeight;
                                rowHeights[(int)row.RowIndex.Value - 1] = rowHeight;

                                writer.Write($"\n{new string(' ', 12)}<tr>");

                                List<Cell> cells = new List<Cell>();

                                for (int i = 0; i < totalColumn; i++)
                                {
                                    cells.Add(new Cell() { CellValue = new CellValue(""), CellReference = ((char)(64 + i + 1)).ToString() + row.RowIndex });
                                }
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    cells[(int)GetColumnIndex(GetColumnName(cell.CellReference))] = cell;
                                }

                                foreach (Cell cell in cells)
                                {
                                    int addedColumnNumber = 1;

                                    uint columnSpanned = 1;
                                    uint rowSpanned = 1;

                                    double actualCellHeight = config.ConvertSize ? rowHeight : double.NaN;
                                    double actualCellWidth = config.ConvertSize ? (currentColumn >= columnWidths.Count ? double.NaN : columnWidths[currentColumn]) : double.NaN;

                                    if (containsMergeCells && cell.CellReference != null)
                                    {
                                        string columnName = GetColumnName(cell.CellReference);
                                        uint columnIndex = GetColumnIndex(columnName);
                                        uint rowIndex = GetRowIndex(cell.CellReference);

                                        if (mergeCells.Any(x => (rowIndex == x.FromRow && columnIndex == x.FromColumn) == false && rowIndex >= x.FromRow && rowIndex <= x.ToRow && columnIndex >= x.FromColumn && columnIndex <= x.ToColumn))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            foreach (MergeCellInfo mergeCellInfo in mergeCells)
                                            {
                                                if (GetColumnIndex(columnName) == mergeCellInfo.FromColumn && rowIndex == mergeCellInfo.FromRow)
                                                {
                                                    addedColumnNumber = (int)mergeCellInfo.ColumnSpanned;

                                                    columnSpanned = mergeCellInfo.ColumnSpanned;
                                                    rowSpanned = mergeCellInfo.RowSpanned;
                                                    actualCellWidth = columnSpanned > 1 ? double.NaN : actualCellWidth;
                                                    actualCellHeight = rowSpanned > 1 ? double.NaN : actualCellHeight;

                                                    break;
                                                }
                                            }
                                        }
                                    }

                                    string cellValue = "";
                                    if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null && int.TryParse(cell.CellValue.Text, out int sharedStringId))
                                    {
                                        SharedStringItem sharedString = (SharedStringItem)sharedStringTable.ChildElements[sharedStringId];

                                        try
                                        {
                                            Run lastRun = null;

                                            foreach (OpenXmlElement element in sharedString.Descendants())
                                            {
                                                if (element is Text text && (lastRun == null || lastRun.Text != text))
                                                {
                                                    cellValue += text.Text;
                                                }
                                                else if (element is Run run)
                                                {
                                                    lastRun = run;

                                                    string runStyle = "";
                                                    if (config.ConvertStyle)
                                                    {
                                                        if (run.RunProperties.GetFirstChild<Color>() is Color fontColor)
                                                        {
                                                            string value = GetColorFromColorType(document, fontColor, new RgbaColor() { R = 0, G = 0, B = 0, A = 1 });
                                                            runStyle += $" color: {value};";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<Bold>() is Bold bold)
                                                        {
                                                            bool value = bold.Val == null || bold.Val.Value;
                                                            runStyle += $" font-weight: {(value ? "bold" : "normal")};";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<Italic>() is Italic italic)
                                                        {
                                                            bool value = italic.Val == null || italic.Val.Value;
                                                            runStyle += $" font-style: {(value ? "italic" : "normal")};";
                                                        }
                                                        string textDecoraion = "";
                                                        if (run.RunProperties.GetFirstChild<Strike>() is Strike strike)
                                                        {
                                                            bool value = strike.Val == null || strike.Val.Value;
                                                            textDecoraion += value ? " line-through" : " none";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<Underline>() is Underline underline)
                                                        {
                                                            if (underline.Val != null)
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
                                                                        textDecoraion += " underline";
                                                                        break;
                                                                    case UnderlineValues.DoubleAccounting:
                                                                        textDecoraion += " underline";
                                                                        break;
                                                                }
                                                            }
                                                        }
                                                        runStyle += $" text-decoration:{textDecoraion};";
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
                                                    }

                                                    cellValue += $"<p style=\"display: inline; {runStyle}\">{(run.Text.Text ?? run.Text.InnerText)}</p>";
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
                                        if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Date)
                                        {
                                            try
                                            {
                                                DateTime dateValue = DateTime.FromOADate(double.Parse(cell.CellValue.Text)).Date;

                                                if (cell.StyleIndex != null && styles.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value] is CellFormat cellFormat)
                                                {
                                                    try
                                                    {
                                                        if (styles.Stylesheet.NumberingFormats != null)
                                                        {
                                                            NumberingFormat numberingFormat = styles.Stylesheet.NumberingFormats.ChildElements.First(x => ((NumberingFormat)x).NumberFormatId == cellFormat.NumberFormatId.Value) as NumberingFormat;

                                                            if (cellFormat.NumberFormatId.Value != 0 && numberingFormat.FormatCode.Value != "@")
                                                            {
                                                                string format = numberingFormat.FormatCode.Value.Replace("&quot;", "");

                                                                if (format.ToLower() == "d")
                                                                {
                                                                    cellValue = dateValue.Day.ToString();
                                                                }
                                                                else if (format.ToLower() == "m")
                                                                {
                                                                    cellValue = dateValue.Month.ToString();
                                                                }
                                                                else if (format.ToLower() == "y")
                                                                {
                                                                    cellValue = dateValue.Year.ToString();
                                                                }
                                                                else
                                                                {
                                                                    cellValue = dateValue.ToString(format);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                cellValue = dateValue.ToString();
                                                            }
                                                        }
                                                        else
                                                        {
                                                            cellValue = cell.CellValue.Text;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        cellValue = dateValue.ToString();
                                                    }
                                                }
                                                else
                                                {
                                                    cellValue = dateValue.ToString();
                                                }
                                            }
                                            catch
                                            {
                                                cellValue = cell.CellValue.Text;
                                            }
                                        }
                                        else
                                        {
                                            cellValue = cell.CellValue.Text;

                                            if (cell.StyleIndex != null && styles.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value] is CellFormat cellFormat)
                                            {
                                                try
                                                {
                                                    if (styles.Stylesheet.NumberingFormats != null)
                                                    {
                                                        NumberingFormat numberingFormat = styles.Stylesheet.NumberingFormats.ChildElements.First(x => ((NumberingFormat)x).NumberFormatId == cellFormat.NumberFormatId.Value) as NumberingFormat;

                                                        if (cellFormat.NumberFormatId.Value != 0 && numberingFormat.FormatCode.Value != "@")
                                                        {
                                                            cellValue = string.Format($"{{0:{numberingFormat.FormatCode.Value}}}", cellValue);
                                                        }
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                    }

                                    string styleHtml = "";
                                    if (config.ConvertStyle)
                                    {
                                        int differentialStyleIndex = -1;
                                        int styleIndex = (cell.StyleIndex != null && cell.StyleIndex.HasValue) ? (int)cell.StyleIndex.Value : ((row.StyleIndex != null && row.StyleIndex.HasValue) ? (int)row.StyleIndex.Value : -1);

                                        if (conditionalFormatting != null && conditionalFormatting.SequenceOfReferences != null && conditionalFormatting.SequenceOfReferences.HasValue && conditionalFormatting.SequenceOfReferences.Any(x => x.HasValue ? x.Value == cell.CellReference : x.InnerText == cell.CellReference))
                                        {
                                            int currentPriority = -1;

                                            try
                                            {
                                                foreach (ConditionalFormattingRule rule in conditionalFormatting.Elements<ConditionalFormattingRule>())
                                                {
                                                    try
                                                    {
                                                        if (rule.Priority == null || !rule.Priority.HasValue || rule.FormatId == null || !rule.FormatId.HasValue || (currentPriority != -1 && rule.Priority.Value > currentPriority))
                                                        {
                                                            continue;
                                                        }

                                                        if (rule.Type != null && rule.GetFirstChild<Formula>() is Formula formula)
                                                        {
                                                            ConditionalFormattingOperatorValues ruleOperator = rule.Operator != null && rule.Operator.HasValue ? rule.Operator.Value : ConditionalFormattingOperatorValues.Equal;
                                                            string formulaText = formula.Text.TrimStart('"').TrimEnd('"');

                                                            if (rule.Type == ConditionalFormatValues.CellIs)
                                                            {
                                                                switch (ruleOperator)
                                                                {
                                                                    case ConditionalFormattingOperatorValues.Equal:
                                                                        if (cellValue.Equals(formulaText))
                                                                        {
                                                                            differentialStyleIndex = (int)rule.FormatId.Value;
                                                                        }
                                                                        break;
                                                                    case ConditionalFormattingOperatorValues.BeginsWith:
                                                                        if (cellValue.StartsWith(formulaText))
                                                                        {
                                                                            differentialStyleIndex = (int)rule.FormatId.Value;
                                                                        }
                                                                        break;
                                                                    case ConditionalFormattingOperatorValues.EndsWith:
                                                                        if (cellValue.EndsWith(formulaText))
                                                                        {
                                                                            differentialStyleIndex = (int)rule.FormatId.Value;
                                                                        }
                                                                        break;
                                                                    case ConditionalFormattingOperatorValues.ContainsText:
                                                                        if (cellValue.Contains(formulaText))
                                                                        {
                                                                            differentialStyleIndex = (int)rule.FormatId.Value;
                                                                        }
                                                                        break;
                                                                    default:
                                                                        continue;
                                                                }
                                                            }
                                                            else if (rule.Type == ConditionalFormatValues.BeginsWith)
                                                            {
                                                                if (cellValue.StartsWith(formulaText))
                                                                {
                                                                    differentialStyleIndex = (int)rule.FormatId.Value;
                                                                }
                                                            }
                                                            else if (rule.Type == ConditionalFormatValues.EndsWith)
                                                            {
                                                                if (cellValue.EndsWith(formulaText))
                                                                {
                                                                    differentialStyleIndex = (int)rule.FormatId.Value;
                                                                }
                                                            }
                                                        }

                                                        currentPriority = rule.Priority.Value;
                                                    }
                                                    catch
                                                    {
                                                        continue;
                                                    }
                                                }
                                            }
                                            catch { }
                                        }

                                        if (styleIndex != -1 && styles.Stylesheet.CellFormats.ChildElements[styleIndex] is CellFormat cellFormat)
                                        {
                                            try
                                            {
                                                Fill fill = cellFormat.FillId.HasValue ? (Fill)styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value] : null;
                                                Font font = cellFormat.FontId.HasValue ? (Font)styles.Stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value] : null;
                                                Border border = cellFormat.BorderId.HasValue ? (Border)styles.Stylesheet.Borders.ChildElements[(int)cellFormat.BorderId.Value] : null;
                                                styleHtml += GetStyleFromCellFormat(document, styles, cell, fill, font, border, cellFormat.Alignment, ref cellValue);
                                            }
                                            catch { }
                                        }
                                        if (differentialStyleIndex != -1 && styles.Stylesheet.DifferentialFormats.ChildElements[differentialStyleIndex] is DifferentialFormat differentialCellFormat)
                                        {
                                            try
                                            {
                                                styleHtml += GetStyleFromCellFormat(document, styles, cell, differentialCellFormat.Fill, differentialCellFormat.Font, differentialCellFormat.Border, differentialCellFormat.Alignment, ref cellValue);
                                            }
                                            catch { }
                                        }
                                    }

                                    writer.Write($"\n{new string(' ', 16)}<td colspan=\"{columnSpanned}\" rowspan=\"{rowSpanned}\" style=\"height: {(double.IsNaN(actualCellHeight) ? "auto" : actualCellHeight + "px")}; width: {(double.IsNaN(actualCellWidth) ? "auto" : actualCellWidth + "px")};{styleHtml}\">{cellValue}</td>");

                                    currentColumn += addedColumnNumber;
                                }

                                writer.Write($"\n{new string(' ', 12)}</tr>");

                                lastRow = row.RowIndex.Value;

                                progressCallback?.Invoke(document, new ConverterProgressCallbackEventArgs(progressSheetIndex, totalSheets, progressRowIndex, (int)totalRows));
                            }

                            writer.Write($"\n{new string(' ', 8)}</table>");

                            if (worksheet.DrawingsPart != null && worksheet.DrawingsPart.WorksheetDrawing != null)
                            {
                                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absoluteAnchor in worksheet.DrawingsPart.WorksheetDrawing.OfType<DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor>())
                                {
                                    try
                                    {
                                        double left = double.NaN;
                                        double top = double.NaN;
                                        double width = double.NaN;
                                        double height = double.NaN;

                                        if (absoluteAnchor != null)
                                        {
                                            left = absoluteAnchor.Position != null && absoluteAnchor.Position.X != null ? (double)absoluteAnchor.Position.X.Value / 2 * 96 / 72 : double.NaN;
                                            top = absoluteAnchor.Position != null && absoluteAnchor.Position.Y != null ? (double)absoluteAnchor.Position.Y.Value / 2 * 96 / 72 : double.NaN;
                                            width = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cx != null ? (double)absoluteAnchor.Extent.Cx.Value / 914400 * 96 : double.NaN;
                                            height = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cy != null == false ? (double)absoluteAnchor.Extent.Cy.Value / 914400 * 96 : double.NaN;
                                        }

                                        ConvertDrawings(writer, config.ConvertPicture, worksheet, absoluteAnchor, left, top, width, height);
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }

                                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>())
                                {
                                    try
                                    {
                                        double left = double.NaN;
                                        double top = double.NaN;
                                        double width = double.NaN;
                                        double height = double.NaN;

                                        if (oneCellAnchor != null)
                                        {
                                            left = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.ColumnId != null ? columnWidths.Take(int.Parse(oneCellAnchor.FromMarker.ColumnId.Text)).Sum() + (oneCellAnchor.FromMarker.ColumnOffset != null ? double.Parse(oneCellAnchor.FromMarker.ColumnOffset.Text) / 914400 * 96 : 0) : double.NaN;
                                            top = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.RowId != null ? rowHeights.Take(int.Parse(oneCellAnchor.FromMarker.RowId.Text)).Sum() + (oneCellAnchor.FromMarker.RowOffset != null ? double.Parse(oneCellAnchor.FromMarker.RowOffset.Text) / 914400 * 96 : 0) : double.NaN;
                                            width = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cx != null ? (double)oneCellAnchor.Extent.Cx.Value / 914400 * 96 : double.NaN;
                                            height = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cy != null == false ? (double)oneCellAnchor.Extent.Cy.Value / 914400 * 96 : double.NaN;
                                        }
                                        ConvertDrawings(writer, config.ConvertPicture, worksheet, oneCellAnchor, left, top, width, height);
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }

                                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                                {
                                    try
                                    {
                                        double left = double.NaN;
                                        double top = double.NaN;
                                        double width = double.NaN;
                                        double height = double.NaN;

                                        if (twoCellAnchor != null)
                                        {
                                            double fromLeft = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null ? columnWidths.Take(int.Parse(twoCellAnchor.FromMarker.ColumnId.Text)).Sum() + (twoCellAnchor.FromMarker.ColumnOffset != null ? double.Parse(twoCellAnchor.FromMarker.ColumnOffset.Text) / 914400 * 96 : 0) : double.NaN;
                                            double fromTop = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null ? rowHeights.Take(int.Parse(twoCellAnchor.FromMarker.RowId.Text)).Sum() + (twoCellAnchor.FromMarker.RowOffset != null ? double.Parse(twoCellAnchor.FromMarker.RowOffset.Text) / 914400 * 96 : 0) : double.NaN;
                                            double toLeft = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null ? columnWidths.Take(int.Parse(twoCellAnchor.ToMarker.ColumnId.Text)).Sum() + (twoCellAnchor.ToMarker.ColumnOffset != null ? double.Parse(twoCellAnchor.ToMarker.ColumnOffset.Text) / 914400 * 96 : 0) : double.NaN;
                                            double toTop = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null ? rowHeights.Take(int.Parse(twoCellAnchor.ToMarker.RowId.Text)).Sum() + (twoCellAnchor.ToMarker.RowOffset != null ? double.Parse(twoCellAnchor.ToMarker.RowOffset.Text) / 914400 * 96 : 0) : double.NaN;

                                            left = Math.Min(fromLeft, toLeft);
                                            top = Math.Min(fromTop, toTop);
                                            width = Math.Abs(toLeft - fromLeft);
                                            height = Math.Abs(toTop - fromTop);
                                        }

                                        ConvertDrawings(writer, config.ConvertPicture, worksheet, twoCellAnchor, left, top, width, height);
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }

                            writer.Write($"\n{new string(' ', 4)}</div>");
                        }
                    }

                    if (config.ConvertHtmlBodyOnly == false)
                    {
                        writer.Write("\n</body>\n</html>");
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                System.Diagnostics.Debug.WriteLine("XlsxToHtmlConverter exception (exceptions only display in Debug mode): " + ex.Message);
#endif

                    outputHtml.Seek(0, SeekOrigin.Begin);
                    outputHtml.SetLength(0);
                    writer.Write(config.ErrorMessage.Replace("{EXCEPTION}", ex.Message));
                }
                finally
                {
                    outputHtml.Seek(0, SeekOrigin.Begin);
                }
            }
        }

        #endregion

        #region Private Methods

        private static uint GetColumnIndex(string columnName)
        {
            columnName = columnName.ToUpper();

            int columnNumber = -1;
            int mulitplier = 1;

            foreach (char c in columnName.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * (c - 64);
                mulitplier *= 26;
            }

            return (uint)Math.Max(0, columnNumber);
        }

        private static string GetColumnName(string cellName)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[A-Za-z]+");
            System.Text.RegularExpressions.Match match = regex.Match(cellName);

            return match.Value;
        }

        private static uint GetRowIndex(string cellName)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"\d+");
            System.Text.RegularExpressions.Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        private static void BorderStyleToHtmlAttributes(SpreadsheetDocument document, BorderPropertiesType border, ref string width, ref string style, ref string color)
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

            color = border.Color != null ? GetColorFromColorType(document, border.Color, new RgbaColor() { R = 211, G = 211, B = 211, A = 1 }) : "lightgray";
        }

        private static string GetColorFromColorType(SpreadsheetDocument document, ColorType type, RgbaColor defaultColor, ColorType background = null)
        {
            RgbaColor rgbColor = new RgbaColor() { R = 0, G = 0, B = 0, A = 1 };

            try
            {
                if (type == null)
                {
                    throw new Exception("Doesn't have any color. Please use default value.");
                }

                if (type.Rgb != null)
                {
                    rgbColor = HexToRgba(type.Rgb.Value);
                }
                else if (type.Indexed != null)
                {
                    if (type.Indexed.Value < IndexedColorData.Length)
                    {
                        rgbColor = HexToRgba(IndexedColorData[type.Indexed.Value]);
                    }
                    else
                    {
                        throw new Exception("Doesn't have any color value from index. Please use default value.");
                    }
                }
                else if (type.Theme != null)
                {
                    DocumentFormat.OpenXml.Drawing.Color2Type color = (DocumentFormat.OpenXml.Drawing.Color2Type)document.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)type.Theme.Value + 2];

                    if (color.RgbColorModelHex != null)
                    {
                        rgbColor = HexToRgba(color.RgbColorModelHex.Val.Value);
                    }
                    else if (color.RgbColorModelPercentage != null)
                    {
                        rgbColor.R = color.RgbColorModelPercentage.RedPortion / 1000 / 100 * 255;
                        rgbColor.G = color.RgbColorModelPercentage.GreenPortion / 1000 / 100 * 255;
                        rgbColor.B = color.RgbColorModelPercentage.BluePortion / 1000 / 100 * 255;
                    }
                    else if (color.HslColor != null)
                    {
                        HlsToRgb(color.HslColor.HueValue, color.HslColor.LumValue, color.HslColor.SatValue, out int r, out int g, out int b);
                        rgbColor.R = r;
                        rgbColor.G = g;
                        rgbColor.B = b;
                    }
                    else if (color.SystemColor != null)
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
                                throw new Exception("Doesn't have any custom value. Please use default value.");
                        };
                    }
                    else if (color.PresetColor != null)
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
                                throw new Exception("Doesn't have any custom value. Please use default value.");
                        };
                    }
                    else
                    {
                        throw new Exception("Doesn't have any custom value. Please use default value.");
                    }
                }
                else if (background != null)
                {
                    try
                    {
                        RgbaColor white = new RgbaColor() { R = 255, G = 255, B = 255, A = 1 };
                        string[] colors = GetColorFromColorType(document, background, white).Replace("rgb(", "").Replace("rgba(", "").Replace(")", "").Split(',');

                        if (colors.Length == 4)
                        {
                            rgbColor.R = int.Parse(colors[0].Trim());
                            rgbColor.G = int.Parse(colors[1].Trim());
                            rgbColor.B = int.Parse(colors[2].Trim());
                            rgbColor.A = double.Parse(colors[3].Trim());
                        }
                        else
                        {
                            rgbColor.R = int.Parse(colors[0].Trim());
                            rgbColor.G = int.Parse(colors[1].Trim());
                            rgbColor.B = int.Parse(colors[2].Trim());
                        }
                    }
                    catch
                    {
                        throw new Exception("Can't get the background color. Please use default value.");
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

                    int r = rgbColor.R;
                    int g = rgbColor.G;
                    int b = rgbColor.B;
                    RgbToHls(r, g, b, out double h, out double l, out double s);

                    if (tint < 0)
                    {
                        HlsToRgb(h, l * (1 + tint), s, out r, out g, out b);
                    }
                    else
                    {
                        HlsToRgb(h, l * (1 - tint) + 1 - 1 * (1 - tint), s, out r, out g, out b);
                    }

                    rgbColor.R = r;
                    rgbColor.G = g;
                    rgbColor.B = b;
                }
            }
            catch { }

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
            hex = hex.Replace("#", "");
            return new RgbaColor()
            {
                A = hex.Length == 8 ? Convert.ToInt32(hex.Substring(0, 2), 16) / 255.0 : 1,
                R = Convert.ToInt32(hex.Substring(hex.Length == 8 ? 2 : 0, 2), 16),
                G = Convert.ToInt32(hex.Substring(hex.Length == 8 ? 4 : 2, 2), 16),
                B = Convert.ToInt32(hex.Substring(hex.Length == 8 ? 6 : 4, 2), 16)
            };
        }

        private static void RgbToHls(int r, int g, int b, out double h, out double l, out double s)
        {
            double double_r = r / 255.0;
            double double_g = g / 255.0;
            double double_b = b / 255.0;

            double max = Math.Max(double_r, Math.Max(double_g, double_b));
            double min = Math.Min(double_r, Math.Min(double_g, double_b));

            double diff = max - min;
            l = (max + min) / 2;
            if (Math.Abs(diff) < 0.00001)
            {
                s = 0;
                h = 0;
            }
            else
            {
                s = l <= 0.5 ? diff / (max + min) : diff / (2 - max - min);

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
                h += h < 0 ? 360 : 0;
            }
        }

        private static void HlsToRgb(double h, double l, double s, out int r, out int g, out int b)
        {
            double p2 = l <= 0.5 ? l * (1 + s) : l + s - l * s;
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

        private static string GetStyleFromCellFormat(SpreadsheetDocument document, WorkbookStylesPart styles, Cell cell, Fill fill, Font font, Border border, Alignment alignment, ref string cellValue)
        {
            string styleHtml = "";

            try
            {
                if (fill.PatternFill != null && fill.PatternFill.PatternType.Value != PatternValues.None)
                {
                    string background = "transparent";
                    if (fill.PatternFill.ForegroundColor != null)
                    {
                        background = GetColorFromColorType(document, fill.PatternFill.ForegroundColor, new RgbaColor() { R = 255, G = 255, B = 255, A = 1 }, fill.PatternFill.BackgroundColor ?? null);
                    }
                    else if (fill.PatternFill.BackgroundColor != null)
                    {
                        background = GetColorFromColorType(document, fill.PatternFill.BackgroundColor, new RgbaColor() { R = 255, G = 255, B = 255, A = 1 });
                    }
                    styleHtml += $" background-color: {background};";
                }
            }
            catch { }
            try
            {
                string fontColor = font.Color != null ? GetColorFromColorType(document, font.Color, new RgbaColor() { R = 0, G = 0, B = 0, A = 1 }) : "black";
                double fontSize = font.FontSize != null && font.FontSize.Val != null ? font.FontSize.Val.Value * 96 / 72 : 11;
                string fontFamily = font.FontName != null && font.FontName.Val != null ? $"\'{font.FontName.Val.Value}\', serif" : "serif";
                bool bold = font.Bold != null && font.Bold.Val != null ? font.Bold.Val.Value : (font.Bold != null);
                bool italic = font.Italic != null && font.Italic.Val != null ? font.Italic.Val.Value : (font.Italic != null);
                bool strike = font.Strike != null && font.Strike.Val != null ? font.Strike.Val.Value : (font.Strike != null);
                string underline = "";
                if (font.Underline != null && font.Underline.Val != null)
                {
                    switch (font.Underline.Val.Value)
                    {
                        case UnderlineValues.Single:
                            underline = " underline";
                            break;
                        case UnderlineValues.SingleAccounting:
                            underline = " underline";
                            break;
                        case UnderlineValues.Double:
                            underline = " underline";
                            break;
                        case UnderlineValues.DoubleAccounting:
                            underline = " underline";
                            break;
                    }
                }

                styleHtml += $" color: {fontColor}; font-size: {fontSize}px; font-family: {fontFamily}; font-weight: {(bold ? "bold" : "normal")}; font-style: {(italic ? "italic" : "normal")}; text-decoration: {(strike ? "line-through" : "none")}{underline};";
            }
            catch { }
            try
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
                    BorderStyleToHtmlAttributes(document, leftBorder, ref leftWidth, ref leftStyle, ref leftColor);
                }
                if (border.RightBorder is RightBorder rightBorder)
                {
                    BorderStyleToHtmlAttributes(document, rightBorder, ref rightWidth, ref rightStyle, ref rightColor);
                }
                if (border.TopBorder is TopBorder topBorder)
                {
                    BorderStyleToHtmlAttributes(document, topBorder, ref topWidth, ref topStyle, ref topColor);
                }
                if (border.BottomBorder is BottomBorder bottomBorder)
                {
                    BorderStyleToHtmlAttributes(document, bottomBorder, ref bottomWidth, ref bottomStyle, ref bottomColor);
                }

                styleHtml += $" border-width: {topWidth} {rightWidth} {bottomWidth} {leftWidth}; border-style: {topStyle} {rightStyle} {bottomStyle} {leftStyle}; border-color: {topColor} {rightColor} {bottomColor} {leftColor};";
            }
            catch { }
            try
            {
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
                            case HorizontalAlignmentValues.General:
                                horizontalTextAlignment = cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Number ? "right" : "left";
                                break;
                        }
                    }

                    styleHtml += $" text-align: {horizontalTextAlignment}; vertical-align: {verticalTextAlignment};";

                    if (alignment.WrapText != null && alignment.WrapText.Value)
                    {
                        styleHtml += " word-wrap: break-word; white-space: normal;";
                    }
                    if (alignment.TextRotation != null)
                    {
                        cellValue = $"<div style=\"width: fit-content; transform: rotate(-{alignment.TextRotation.Value}deg);\">" + cellValue + "</div>";
                    }
                }
            }
            catch { }

            return styleHtml;
        }

        private static void ConvertDrawings(StreamWriter writer, bool convertPicture, WorksheetPart worksheet, OpenXmlCompositeElement anchor, double left, double top, double width, double height)
        {
            if (convertPicture)
            {
                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>())
                {
                    try
                    {
                        if (picture.NonVisualPictureProperties != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Hidden != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Hidden.HasValue && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Hidden.Value)
                        {
                            continue;
                        }

                        int rotation = 0;
                        double pictureLeft = left;
                        double pictureTop = top;
                        double pictureWidth = width;
                        double pictureHeight = height;
                        string alt = picture.NonVisualPictureProperties != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description != null && picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.HasValue ? picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.Value : "Image";

                        if (picture.ShapeProperties != null && picture.ShapeProperties.Transform2D != null)
                        {
                            if (picture.ShapeProperties.Transform2D.Offset != null)
                            {
                                if (!double.IsNaN(pictureLeft))
                                {
                                    pictureLeft += picture.ShapeProperties.Transform2D.Offset.X != null && picture.ShapeProperties.Transform2D.Offset.X.HasValue ? picture.ShapeProperties.Transform2D.Offset.X.Value / 914400 * 96 : 0;
                                }
                                if (!double.IsNaN(pictureTop))
                                {
                                    pictureTop += picture.ShapeProperties.Transform2D.Offset.Y != null && picture.ShapeProperties.Transform2D.Offset.Y.HasValue ? picture.ShapeProperties.Transform2D.Offset.Y.Value / 914400 * 96 : 0;
                                }
                            }
                            if (picture.ShapeProperties.Transform2D.Extents != null)
                            {
                                if (!double.IsNaN(pictureWidth))
                                {
                                    pictureWidth += picture.ShapeProperties.Transform2D.Extents.Cx != null && picture.ShapeProperties.Transform2D.Extents.Cx.HasValue ? picture.ShapeProperties.Transform2D.Extents.Cx.Value / 914400 * 96 : 0;
                                }
                                if (!double.IsNaN(pictureHeight))
                                {
                                    pictureHeight += picture.ShapeProperties.Transform2D.Extents.Cy != null && picture.ShapeProperties.Transform2D.Extents.Cy.HasValue ? picture.ShapeProperties.Transform2D.Extents.Cy.Value / 914400 * 96 : 0;
                                }
                            }
                            if (picture.ShapeProperties.Transform2D.Rotation != null && picture.ShapeProperties.Transform2D.Rotation.HasValue)
                            {
                                rotation = picture.ShapeProperties.Transform2D.Rotation.Value;
                            }
                        }

                        ImagePart imagePart = worksheet.DrawingsPart.GetPartById(picture.BlipFill.Blip.Embed.Value) as ImagePart;

                        Stream imageStream = imagePart.GetStream();
                        imageStream.Seek(0, SeekOrigin.Begin);
                        byte[] data = new byte[imageStream.Length];
                        imageStream.Read(data, 0, (int)imageStream.Length);
                        string base64 = Convert.ToBase64String(data, Base64FormattingOptions.None);

                        writer.Write($"\n{new string(' ', 8)}<img alt=\"{alt}\" src=\"data:{imagePart.ContentType};base64,{base64}\" style=\"position: absolute; left: {(double.IsNaN(pictureLeft) == false ? pictureLeft + "px" : "0px")}; top: {(double.IsNaN(pictureTop) == false ? pictureTop + "px" : "0px")}; width: {(double.IsNaN(pictureWidth) == false ? pictureWidth + "px" : "auto")}; height: {(double.IsNaN(pictureHeight) == false ? pictureHeight + "px" : "auto")}; {(rotation != 0 ? $"transform: rotate(-{rotation}deg);" : "")}\"/>");
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>())
            {
                try
                {
                    if (shape.NonVisualShapeProperties != null && shape.NonVisualShapeProperties.NonVisualDrawingProperties != null && shape.NonVisualShapeProperties.NonVisualDrawingProperties.Hidden != null && shape.NonVisualShapeProperties.NonVisualDrawingProperties.Hidden.HasValue && shape.NonVisualShapeProperties.NonVisualDrawingProperties.Hidden.Value)
                    {
                        continue;
                    }

                    int rotation = 0;
                    double shapeLeft = left;
                    double shapeTop = top;
                    double shapeWidth = width;
                    double shapeHeight = height;

                    if (shape.ShapeProperties != null && shape.ShapeProperties.Transform2D != null)
                    {
                        if (shape.ShapeProperties.Transform2D.Offset != null)
                        {
                            if (!double.IsNaN(shapeLeft))
                            {
                                shapeLeft += shape.ShapeProperties.Transform2D.Offset.X != null && shape.ShapeProperties.Transform2D.Offset.X.HasValue ? shape.ShapeProperties.Transform2D.Offset.X.Value / 914400 * 96 : 0;
                            }
                            if (!double.IsNaN(shapeTop))
                            {
                                shapeTop += shape.ShapeProperties.Transform2D.Offset.Y != null && shape.ShapeProperties.Transform2D.Offset.Y.HasValue ? shape.ShapeProperties.Transform2D.Offset.Y.Value / 914400 * 96 : 0;
                            }
                        }
                        if (shape.ShapeProperties.Transform2D.Extents != null)
                        {
                            if (!double.IsNaN(shapeWidth))
                            {
                                shapeWidth += shape.ShapeProperties.Transform2D.Extents.Cx != null && shape.ShapeProperties.Transform2D.Extents.Cx.HasValue ? shape.ShapeProperties.Transform2D.Extents.Cx.Value / 914400 * 96 : 0;
                            }
                            if (!double.IsNaN(shapeHeight))
                            {
                                shapeHeight += shape.ShapeProperties.Transform2D.Extents.Cy != null && shape.ShapeProperties.Transform2D.Extents.Cy.HasValue ? shape.ShapeProperties.Transform2D.Extents.Cy.Value / 914400 * 96 : 0;
                            }
                        }
                        if (shape.ShapeProperties.Transform2D.Rotation != null && shape.ShapeProperties.Transform2D.Rotation.HasValue)
                        {
                            rotation = shape.ShapeProperties.Transform2D.Rotation.Value;
                        }
                    }
                    string text = shape.TextBody != null ? shape.TextBody.InnerText : "";

                    writer.Write($"\n{new string(' ', 8)}<p style=\"position: absolute; left: {(double.IsNaN(shapeLeft) == false ? shapeLeft + "px" : "0px")}; top: {(double.IsNaN(shapeTop) == false ? shapeTop + "px" : "0px")}; width: {(double.IsNaN(shapeWidth) == false ? shapeWidth + "px" : "auto")}; height: {(double.IsNaN(shapeHeight) == false ? shapeHeight + "px" : "auto")}; {(rotation != 0 ? $"transform: rotate(-{rotation}deg);" : "")}\">{text}</p>");
                }
                catch
                {
                    continue;
                }
            }
        }

        #endregion

        #region Private Structure

        private struct MergeCellInfo
        {
            public uint FromColumn { get; set; }
            public uint FromRow { get; set; }
            public uint ToColumn { get; set; }
            public uint ToRow { get; set; }
            public uint ColumnSpanned { get; set; }
            public uint RowSpanned { get; set; }
        }

        private struct RgbaColor
        {
            public int R { get; set; }
            public int G { get; set; }
            public int B { get; set; }
            public double A { get; set; }
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
