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
            config = config ?? ConverterConfig.DefaultSettings;

            outputHtml.Seek(0, SeekOrigin.Begin);
            outputHtml.SetLength(0);

            using (StreamWriter writer = new StreamWriter(outputHtml, config.Encoding, 65536))
            {
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
                        WorkbookStylesPart styles = workbook.WorkbookStylesPart;
                        IEnumerable<Sheet> sheets = workbook.Workbook.Descendants<Sheet>();

                        SharedStringTable sharedStringTable = null;
                        IEnumerable<SharedStringTablePart> sharedStringTables = workbook.GetPartsOfType<SharedStringTablePart>();
                        if (sharedStringTables.Any())
                        {
                            sharedStringTable = sharedStringTables.First().SharedStringTable;
                        }

                        int totalSheets = sheets.Count();
                        int progressSheetIndex = 0;
                        foreach (Sheet currentSheet in sheets)
                        {
                            progressSheetIndex++;

                            if (config.ConvertFirstSheetOnly == true && progressSheetIndex > 1)
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

                            if (config.ConvertSheetNameTitle == true)
                            {
                                string tabColor = sheet.SheetProperties != null && sheet.SheetProperties.TabColor != null ? GetColorFromColorType(workbook, sheet.SheetProperties.TabColor) : "";
                                writer.Write($"\n{new string(' ', 4)}<h5{(!string.IsNullOrEmpty(tabColor) ? " style=\"border-bottom-color: " + tabColor + ";\"" : "")}>{(currentSheet.Name != null && currentSheet.Name.HasValue ? currentSheet.Name.Value : "Untitled")}</h5>");
                            }

                            writer.Write($"\n{new string(' ', 4)}<div style=\"position: relative;\">");
                            writer.Write($"\n{new string(' ', 8)}<table>");

                            double defaultRowHeight = sheet.SheetFormatProperties != null && sheet.SheetFormatProperties.DefaultRowHeight != null && sheet.SheetFormatProperties.DefaultRowHeight.HasValue ? sheet.SheetFormatProperties.DefaultRowHeight.Value / 0.75 : 20;

                            bool containsMergeCells = false;
                            List<MergeCellInfo> mergeCells = new List<MergeCellInfo>();
                            if (sheet.Descendants<MergeCells>().FirstOrDefault() is MergeCells mergeCellsGroup)
                            {
                                containsMergeCells = true;

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

                                    mergeCells.Add(new MergeCellInfo() { FromColumn = fromColumn, FromRow = fromRow, ToColumn = toColumn, ToRow = toRow, ColumnSpanned = toColumn - fromColumn + 1, RowSpanned = toRow - fromRow + 1 });
                                }
                            }

                            ConditionalFormatting conditionalFormatting = null;
                            if (sheet.Elements<ConditionalFormatting>() is IEnumerable<ConditionalFormatting> conditionalFormattings && conditionalFormattings.Any())
                            {
                                conditionalFormatting = conditionalFormattings.First();
                            }

                            IEnumerable<Row> rows = sheet.Descendants<Row>();

                            int totalRows = 0;
                            int totalColumn = 0;
                            if (sheet.SheetDimension != null && sheet.SheetDimension.Reference != null && sheet.SheetDimension.Reference.HasValue)
                            {
                                string[] dimension = sheet.SheetDimension.Reference.Value.Split(':');
                                int fromColumn = GetColumnIndex(dimension[0]);
                                int toColumn = GetColumnIndex(dimension[1]);
                                int fromRow = GetRowIndex(dimension[0]);
                                int toRow = GetRowIndex(dimension[1]);

                                totalRows = toRow - fromColumn + 1;
                                totalColumn = toColumn - fromColumn + 1;
                            }
                            else
                            {
                                foreach (Row row in rows)
                                {
                                    foreach (Cell cell in row.Descendants<Cell>())
                                    {
                                        if (cell.CellReference == null || !cell.CellReference.HasValue)
                                        {
                                            continue;
                                        }

                                        totalColumn = Math.Max(totalColumn, GetColumnIndex(cell.CellReference.Value) + 1);
                                        totalRows = Math.Max(totalRows, GetRowIndex(cell.CellReference.Value));
                                    }
                                }
                            }

                            int currentColumn = 0;
                            int lastRow = 0;

                            List<double> columnWidths = new List<double>(Enumerable.Repeat(double.NaN, totalColumn));
                            List<double> rowHeights = new List<double>(Enumerable.Repeat(defaultRowHeight, totalRows));
                            if (sheet.GetFirstChild<Columns>() is Columns columnsGroup)
                            {
                                foreach (Column column in columnsGroup.Descendants<Column>())
                                {
                                    if (column == null || column.Min == null || !column.Min.HasValue || column.Max == null || !column.Max.HasValue)
                                    {
                                        continue;
                                    }

                                    for (int i = (int)column.Min.Value; i <= (int)column.Max.Value; i++)
                                    {
                                        if (column.CustomWidth != null && column.CustomWidth.HasValue && column.CustomWidth.Value && column.Width != null && column.Width.HasValue)
                                        {
                                            columnWidths[i - 1] = (column.Width.Value - 1) * 7 + 7;
                                        }
                                    }
                                }
                            }

                            int progressRowIndex = 0;
                            foreach (Row row in rows)
                            {
                                progressRowIndex++;

                                if (row.RowIndex == null || !row.RowIndex.HasValue)
                                {
                                    continue;
                                }

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
                                double rowHeight = (row.CustomHeight == null || (row.CustomHeight.HasValue && row.CustomHeight.Value)) && row.Height != null && row.Height.HasValue ? row.Height.Value / 0.75 : defaultRowHeight;
                                rowHeights[(int)row.RowIndex.Value - 1] = rowHeight;

                                writer.Write($"\n{new string(' ', 12)}<tr>");

                                List<Cell> cells = new List<Cell>();
                                for (int i = 0; i < totalColumn; i++)
                                {
                                    cells.Add(new Cell() { CellValue = new CellValue(""), CellReference = ((char)(64 + i + 1)).ToString() + row.RowIndex });
                                }
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    if (cell.CellReference == null || !cell.CellReference.HasValue)
                                    {
                                        continue;
                                    }

                                    cells[GetColumnIndex(cell.CellReference.Value)] = cell;
                                }

                                foreach (Cell cell in cells)
                                {
                                    int addedColumnNumber = 1;

                                    int columnSpanned = 1;
                                    int rowSpanned = 1;

                                    double actualCellHeight = config.ConvertSize ? rowHeight : double.NaN;
                                    double actualCellWidth = config.ConvertSize ? (currentColumn >= columnWidths.Count ? double.NaN : columnWidths[currentColumn]) : double.NaN;

                                    if (containsMergeCells && cell.CellReference != null)
                                    {
                                        int columnIndex = GetColumnIndex(cell.CellReference);
                                        int rowIndex = GetRowIndex(cell.CellReference);

                                        if (mergeCells.Any(x => !(rowIndex == x.FromRow && columnIndex == x.FromColumn) && rowIndex >= x.FromRow && rowIndex <= x.ToRow && columnIndex >= x.FromColumn && columnIndex <= x.ToColumn))
                                        {
                                            continue;
                                        }

                                        foreach (MergeCellInfo mergeCellInfo in mergeCells)
                                        {
                                            if (columnIndex == mergeCellInfo.FromColumn && rowIndex == mergeCellInfo.FromRow)
                                            {
                                                addedColumnNumber = mergeCellInfo.ColumnSpanned;

                                                columnSpanned = mergeCellInfo.ColumnSpanned;
                                                rowSpanned = mergeCellInfo.RowSpanned;
                                                actualCellWidth = columnSpanned > 1 ? double.NaN : actualCellWidth;
                                                actualCellHeight = rowSpanned > 1 ? double.NaN : actualCellHeight;

                                                break;
                                            }
                                        }
                                    }

                                    string cellValue = "";
                                    if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null && int.TryParse(cell.CellValue.Text, out int sharedStringId) && sharedStringTable.HasChildren && sharedStringId < sharedStringTable.ChildElements.Count && sharedStringTable.ChildElements[sharedStringId] is SharedStringItem sharedString)
                                    {
                                        if (sharedString.HasChildren)
                                        {
                                            Run lastRun = null;

                                            foreach (OpenXmlElement element in sharedString.Descendants())
                                            {
                                                if (element is Text text && (lastRun == null || (lastRun.Text != null && lastRun.Text != text)))
                                                {
                                                    cellValue += text.Text;
                                                }
                                                else if (element is Run run && run.Text != null)
                                                {
                                                    lastRun = run;

                                                    string runStyle = "";
                                                    if (config.ConvertStyle && run.RunProperties != null)
                                                    {
                                                        if (run.RunProperties.GetFirstChild<Color>() is Color fontColor)
                                                        {
                                                            string value = GetColorFromColorType(workbook, fontColor);
                                                            if (!string.IsNullOrEmpty(value))
                                                            {
                                                                runStyle += $" color: {value};";
                                                            }
                                                        }
                                                        if (run.RunProperties.GetFirstChild<FontSize>() is FontSize fontSize && fontSize.Val != null && fontSize.Val.HasValue)
                                                        {
                                                            runStyle += $" font-size: {fontSize.Val.Value * 96 / 72}px;";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<RunFont>() is RunFont runFont && runFont.Val != null && runFont.Val.HasValue)
                                                        {
                                                            runStyle += $" font-family: {runFont.Val.Value};";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<Bold>() is Bold bold)
                                                        {
                                                            bool value = bold.Val == null || (bold.Val.HasValue && bold.Val.Value);
                                                            runStyle += $" font-weight: {(value ? "bold" : "normal")};";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<Italic>() is Italic italic)
                                                        {
                                                            bool value = italic.Val == null || (italic.Val.HasValue && italic.Val.Value);
                                                            runStyle += $" font-style: {(value ? "italic" : "normal")};";
                                                        }
                                                        string textDecoraion = "";
                                                        if (run.RunProperties.GetFirstChild<Strike>() is Strike strike)
                                                        {
                                                            bool value = strike.Val == null || (strike.Val.HasValue && strike.Val.Value);
                                                            textDecoraion += value ? " line-through" : " none";
                                                        }
                                                        if (run.RunProperties.GetFirstChild<Underline>() is Underline underline && underline.Val != null && underline.Val.HasValue)
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
                                                        if (!string.IsNullOrEmpty(textDecoraion))
                                                        {
                                                            runStyle += $" text-decoration:{textDecoraion};";
                                                        }
                                                    }

                                                    cellValue += $"<p style=\"display: inline;{runStyle}\">{(string.IsNullOrEmpty(run.Text.Text) ? run.Text.InnerText : run.Text.Text)}</p>";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            cellValue = sharedString.InnerText;
                                        }
                                    }
                                    else if (cell.CellValue != null)
                                    {
                                        cellValue = cell.CellValue.Text;

                                        if (cell.StyleIndex != null && cell.StyleIndex.HasValue && styles != null && styles.Stylesheet != null && styles.Stylesheet.CellFormats != null && styles.Stylesheet.CellFormats.HasChildren && cell.StyleIndex.Value < styles.Stylesheet.CellFormats.ChildElements.Count && styles.Stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value] is CellFormat cellFormat && cellFormat.NumberFormatId != null && cellFormat.NumberFormatId.HasValue && styles.Stylesheet.NumberingFormats != null && styles.Stylesheet.NumberingFormats.HasChildren && styles.Stylesheet.NumberingFormats.ChildElements.FirstOrDefault(x => x is NumberingFormat item && item.NumberFormatId != null && item.NumberFormatId.HasValue && item.NumberFormatId.Value == cellFormat.NumberFormatId.Value) is NumberingFormat numberingFormat)
                                        {
                                            double cellValueDate = 0;
                                            bool dateFormat = cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Date && double.TryParse(cell.CellValue.Text, out cellValueDate);

                                            if (cellFormat.NumberFormatId.Value != 0 && numberingFormat.FormatCode != null && numberingFormat.FormatCode.HasValue && numberingFormat.FormatCode.Value != "@")
                                            {
                                                if (dateFormat)
                                                {
                                                    DateTime dateValue = DateTime.FromOADate(cellValueDate).Date;
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
                                                    cellValue = string.Format($"{{0:{numberingFormat.FormatCode.Value}}}", cell.CellValue.Text);
                                                }
                                            }
                                            else if (dateFormat)
                                            {
                                                cellValue = DateTime.FromOADate(cellValueDate).Date.ToString();
                                            }
                                        }
                                    }

                                    string styleHtml = "";
                                    if (config.ConvertStyle)
                                    {
                                        int differentialStyleIndex = -1;
                                        int styleIndex = cell.StyleIndex != null && cell.StyleIndex.HasValue ? (int)cell.StyleIndex.Value : (row.StyleIndex != null && row.StyleIndex.HasValue ? (int)row.StyleIndex.Value : -1);

                                        if (cell.CellReference != null && cell.CellReference.HasValue && conditionalFormatting != null && conditionalFormatting.SequenceOfReferences != null && conditionalFormatting.SequenceOfReferences.HasValue && conditionalFormatting.SequenceOfReferences.Any(x => x.HasValue ? x.Value == cell.CellReference.Value : (x.InnerText != null && x.InnerText == cell.CellReference.Value)))
                                        {
                                            int currentPriority = -1;

                                            foreach (ConditionalFormattingRule rule in conditionalFormatting.Elements<ConditionalFormattingRule>())
                                            {
                                                if (rule == null || rule.Priority == null || !rule.Priority.HasValue || rule.FormatId == null || !rule.FormatId.HasValue || (currentPriority != -1 && rule.Priority.Value > currentPriority))
                                                {
                                                    continue;
                                                }

                                                if (rule.Type != null && rule.Type.HasValue && rule.GetFirstChild<Formula>() is Formula formula)
                                                {
                                                    ConditionalFormattingOperatorValues ruleOperator = rule.Operator != null && rule.Operator.HasValue ? rule.Operator.Value : ConditionalFormattingOperatorValues.Equal;
                                                    string formulaText = formula.Text.Trim('"');

                                                    if (rule.Type == ConditionalFormatValues.CellIs)
                                                    {
                                                        switch (ruleOperator)
                                                        {
                                                            case ConditionalFormattingOperatorValues.Equal:
                                                                if (cellValue == formulaText)
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
                                        }

                                        if (styleIndex != -1 && styles != null && styles.Stylesheet != null && styles.Stylesheet.CellFormats != null && styles.Stylesheet.CellFormats.HasChildren && styleIndex < styles.Stylesheet.CellFormats.ChildElements.Count && styles.Stylesheet.CellFormats.ChildElements[styleIndex] is CellFormat cellFormat)
                                        {
                                            Fill fill = cellFormat.FillId != null && cellFormat.FillId.HasValue && styles.Stylesheet.Fills != null && styles.Stylesheet.Fills.HasChildren && cellFormat.FillId.Value < styles.Stylesheet.Fills.ChildElements.Count ? (Fill)styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value] : null;
                                            Font font = cellFormat.FontId != null && cellFormat.FontId.HasValue && styles.Stylesheet.Fonts != null && styles.Stylesheet.Fonts.HasChildren && cellFormat.FontId.Value < styles.Stylesheet.Fonts.ChildElements.Count ? (Font)styles.Stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value] : null;
                                            Border border = cellFormat.BorderId != null && cellFormat.BorderId.HasValue && styles.Stylesheet.Borders != null && styles.Stylesheet.Borders.HasChildren && cellFormat.BorderId.Value < styles.Stylesheet.Borders.ChildElements.Count ? (Border)styles.Stylesheet.Borders.ChildElements[(int)cellFormat.BorderId.Value] : null;
                                            styleHtml += GetStyleFromCellFormat(workbook, cell, fill, font, border, cellFormat.Alignment, ref cellValue);
                                        }
                                        if (differentialStyleIndex != -1 && styles != null && styles.Stylesheet != null && styles.Stylesheet.DifferentialFormats != null && styles.Stylesheet.DifferentialFormats.HasChildren && differentialStyleIndex < styles.Stylesheet.DifferentialFormats.ChildElements.Count && styles.Stylesheet.DifferentialFormats.ChildElements[differentialStyleIndex] is DifferentialFormat differentialCellFormat)
                                        {
                                            styleHtml += GetStyleFromCellFormat(workbook, cell, differentialCellFormat.Fill, differentialCellFormat.Font, differentialCellFormat.Border, differentialCellFormat.Alignment, ref cellValue);
                                        }
                                    }

                                    writer.Write($"\n{new string(' ', 16)}<td colspan=\"{columnSpanned}\" rowspan=\"{rowSpanned}\" style=\"height: {(double.IsNaN(actualCellHeight) ? "auto" : actualCellHeight + "px")}; width: {(double.IsNaN(actualCellWidth) ? "auto" : actualCellWidth + "px")};{styleHtml}\">{cellValue}</td>");

                                    currentColumn += addedColumnNumber;
                                }

                                writer.Write($"\n{new string(' ', 12)}</tr>");

                                lastRow = (int)row.RowIndex.Value;

                                progressCallback?.Invoke(document, new ConverterProgressCallbackEventArgs(progressSheetIndex, totalSheets, progressRowIndex, totalRows));
                            }

                            writer.Write($"\n{new string(' ', 8)}</table>");

                            if (worksheet.DrawingsPart != null && worksheet.DrawingsPart.WorksheetDrawing != null)
                            {
                                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absoluteAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor>())
                                {
                                    if (absoluteAnchor == null)
                                    {
                                        continue;
                                    }

                                    double left = absoluteAnchor.Position != null && absoluteAnchor.Position.X != null && absoluteAnchor.Position.X.HasValue ? (double)absoluteAnchor.Position.X.Value / 2 * 96 / 72 : double.NaN;
                                    double top = absoluteAnchor.Position != null && absoluteAnchor.Position.Y != null && absoluteAnchor.Position.Y.HasValue ? (double)absoluteAnchor.Position.Y.Value / 2 * 96 / 72 : double.NaN;
                                    double width = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cx != null && absoluteAnchor.Extent.Cx.HasValue ? (double)absoluteAnchor.Extent.Cx.Value / 914400 * 96 : double.NaN;
                                    double height = absoluteAnchor.Extent != null && absoluteAnchor.Extent.Cy != null && absoluteAnchor.Extent.Cy.HasValue ? (double)absoluteAnchor.Extent.Cy.Value / 914400 * 96 : double.NaN;

                                    ConvertDrawings(worksheet, absoluteAnchor, writer, left, top, width, height, config.ConvertPicture);
                                }

                                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor oneCellAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>())
                                {
                                    if (oneCellAnchor == null)
                                    {
                                        continue;
                                    }

                                    //TODO: auto-sized columns
                                    double left = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.ColumnId != null && int.TryParse(oneCellAnchor.FromMarker.ColumnId.Text, out int columnId) ? columnWidths.Take(Math.Min(columnWidths.Count, columnId)).Sum() + (oneCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(oneCellAnchor.FromMarker.ColumnOffset.Text, out double columnOffset) ? columnOffset / 914400 * 96 : 0) : double.NaN;
                                    double top = oneCellAnchor.FromMarker != null && oneCellAnchor.FromMarker.RowId != null && int.TryParse(oneCellAnchor.FromMarker.RowId.Text, out int rowId) ? rowHeights.Take(Math.Min(rowHeights.Count, rowId)).Sum() + (oneCellAnchor.FromMarker.RowOffset != null && double.TryParse(oneCellAnchor.FromMarker.RowOffset.Text, out double rowOffset) ? rowOffset / 914400 * 96 : 0) : double.NaN;
                                    double width = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cx != null && oneCellAnchor.Extent.Cx.HasValue ? (double)oneCellAnchor.Extent.Cx.Value / 914400 * 96 : double.NaN;
                                    double height = oneCellAnchor.Extent != null && oneCellAnchor.Extent.Cy != null && oneCellAnchor.Extent.Cy.HasValue ? (double)oneCellAnchor.Extent.Cy.Value / 914400 * 96 : double.NaN;

                                    ConvertDrawings(worksheet, oneCellAnchor, writer, left, top, width, height, config.ConvertPicture);
                                }

                                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor in worksheet.DrawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                                {
                                    if (twoCellAnchor == null)
                                    {
                                        continue;
                                    }
                                    //TODO: auto-sized columns
                                    double fromLeft = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.ColumnId != null && int.TryParse(twoCellAnchor.FromMarker.ColumnId.Text, out int fromColumnId) ? columnWidths.Take(Math.Min(columnWidths.Count, fromColumnId)).Sum() + (twoCellAnchor.FromMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.FromMarker.ColumnOffset.Text, out double fromColumnOffset) ? fromColumnOffset / 914400 * 96 : 0) : double.NaN;
                                    double fromTop = twoCellAnchor.FromMarker != null && twoCellAnchor.FromMarker.RowId != null && int.TryParse(twoCellAnchor.FromMarker.RowId.Text, out int fromRowId) ? rowHeights.Take(Math.Min(rowHeights.Count, fromRowId)).Sum() + (twoCellAnchor.FromMarker.RowOffset != null && double.TryParse(twoCellAnchor.FromMarker.RowOffset.Text, out double fromRowOffset) ? fromRowOffset / 914400 * 96 : 0) : double.NaN;
                                    double toLeft = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.ColumnId != null && int.TryParse(twoCellAnchor.ToMarker.ColumnId.Text, out int toColumnId) ? columnWidths.Take(Math.Min(columnWidths.Count, toColumnId)).Sum() + (twoCellAnchor.ToMarker.ColumnOffset != null && double.TryParse(twoCellAnchor.ToMarker.ColumnOffset.Text, out double toColumnOffset) ? toColumnOffset / 914400 * 96 : 0) : double.NaN;
                                    double toTop = twoCellAnchor.ToMarker != null && twoCellAnchor.ToMarker.RowId != null && int.TryParse(twoCellAnchor.ToMarker.RowId.Text, out int toRowId) ? rowHeights.Take(Math.Min(rowHeights.Count, toRowId)).Sum() + (twoCellAnchor.ToMarker.RowOffset != null && double.TryParse(twoCellAnchor.ToMarker.RowOffset.Text, out double toRowOffset) ? toRowOffset / 914400 * 96 : 0) : double.NaN;

                                    //TODO: flip drawing
                                    double left = Math.Min(fromLeft, toLeft);
                                    double top = Math.Min(fromTop, toTop);
                                    double width = Math.Abs(toLeft - fromLeft);
                                    double height = Math.Abs(toTop - fromTop);

                                    ConvertDrawings(worksheet, twoCellAnchor, writer, left, top, width, height, config.ConvertPicture);
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

        private static int GetColumnIndex(string cellName)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[A-Za-z]+");
            System.Text.RegularExpressions.Match match = regex.Match(cellName);

            if (!match.Success)
            {
                return 0;
            }

            int columnNumber = -1;
            int mulitplier = 1;
            foreach (char c in match.Value.ToUpper().ToCharArray().Reverse())
            {
                columnNumber += mulitplier * (c - 64);
                mulitplier *= 26;
            }

            return Math.Max(0, columnNumber);
        }

        private static int GetRowIndex(string cellName)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"\d+");
            System.Text.RegularExpressions.Match match = regex.Match(cellName);

            return int.TryParse(match.Value, out int rowIndex) ? rowIndex : 0;
        }

        private static string GetColorFromColorType(WorkbookPart workbook, ColorType type)
        {
            if (type == null)
            {
                return "";
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
                        return "";
                }
            }
            else if (type.Theme != null && type.Theme.HasValue && workbook.ThemePart != null && workbook.ThemePart.Theme != null && workbook.ThemePart.Theme.ThemeElements != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme != null && workbook.ThemePart.Theme.ThemeElements.ColorScheme.HasChildren && type.Theme.Value + 2 < workbook.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements.Count && workbook.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)type.Theme.Value + 2] is DocumentFormat.OpenXml.Drawing.Color2Type color)
            {
                if (color.RgbColorModelHex != null && color.RgbColorModelHex.Val != null && color.RgbColorModelHex.Val.HasValue)
                {
                    rgbColor = HexToRgba(color.RgbColorModelHex.Val.Value);
                }
                else if (color.RgbColorModelPercentage != null)
                {
                    rgbColor.R = color.RgbColorModelPercentage.RedPortion.HasValue ? color.RgbColorModelPercentage.RedPortion.Value / 1000 / 100 * 255 : 0;
                    rgbColor.G = color.RgbColorModelPercentage.GreenPortion.HasValue ? color.RgbColorModelPercentage.GreenPortion.Value / 1000 / 100 * 255 : 0;
                    rgbColor.B = color.RgbColorModelPercentage.BluePortion.HasValue ? color.RgbColorModelPercentage.BluePortion.Value / 1000 / 100 * 255 : 0;
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
                            return "";
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
                            return "";
                    };
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
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
            hex = hex.Replace("#", "");

            if (hex.Length < 6)
            {
                return new RgbaColor() { R = 0, G = 0, B = 0, A = 0 };
            }
            else
            {
                return new RgbaColor()
                {
                    R = Convert.ToInt32(hex.Substring(hex.Length == 8 ? 2 : 0, 2), 16),
                    G = Convert.ToInt32(hex.Substring(hex.Length == 8 ? 4 : 2, 2), 16),
                    B = Convert.ToInt32(hex.Substring(hex.Length == 8 ? 6 : 4, 2), 16),
                    A = hex.Length == 8 ? Convert.ToInt32(hex.Substring(0, 2), 16) / 255.0 : 1
                };
            }
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

        private static string GetStyleFromCellFormat(WorkbookPart workbook, Cell cell, Fill fill, Font font, Border border, Alignment alignment, ref string cellValue)
        {
            string styleHtml = "";

            if (fill != null && fill.PatternFill != null && fill.PatternFill.PatternType != null && fill.PatternFill.PatternType.HasValue && fill.PatternFill.PatternType.Value != PatternValues.None)
            {
                string background = "";
                if (fill.PatternFill.ForegroundColor != null)
                {
                    background = GetColorFromColorType(workbook, fill.PatternFill.ForegroundColor);
                }
                if (string.IsNullOrEmpty(background) && fill.PatternFill.BackgroundColor != null)
                {
                    background = GetColorFromColorType(workbook, fill.PatternFill.BackgroundColor);
                }
                if (!string.IsNullOrEmpty(background))
                {
                    styleHtml += $" background-color: {background};";
                }
            }

            if (font != null)
            {
                if (font.Color != null)
                {
                    string fontColor = GetColorFromColorType(workbook, font.Color);
                    if (!string.IsNullOrEmpty(fontColor))
                    {
                        styleHtml += $" color: {fontColor};";
                    }
                }
                if (font.FontSize != null && font.FontSize.Val != null && font.FontSize.Val.HasValue)
                {
                    styleHtml += $" font-size: {font.FontSize.Val.Value * 96 / 72}px;";
                }
                if (font.FontName != null && font.FontName.Val != null && font.FontName.Val.HasValue)
                {
                    styleHtml += $" font-family: \'{font.FontName.Val.Value}\';";
                }
                if (font.Bold != null)
                {
                    bool value = font.Bold.Val == null || (font.Bold.Val.HasValue && font.Bold.Val.Value);
                    styleHtml += $" font-weight: {(value ? "bold" : "normal")};";
                }
                if (font.Italic != null)
                {
                    bool value = font.Italic.Val == null || (font.Italic.Val.HasValue && font.Italic.Val.Value);
                    styleHtml += $" font-style: {(value ? "italic" : "normal")};";
                }
                string textDecoraion = "";
                if (font.Strike != null)
                {
                    bool value = font.Strike.Val == null || (font.Strike.Val.HasValue && font.Strike.Val.Value);
                    textDecoraion += value ? " line-through" : " none";
                }
                if (font.Underline != null && font.Underline.Val != null && font.Underline.Val.HasValue)
                {
                    switch (font.Underline.Val.Value)
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
                if (!string.IsNullOrEmpty(textDecoraion))
                {
                    styleHtml += $" text-decoration:{textDecoraion};";
                }
            }

            if (border != null)
            {
                string leftWidth = "revert";
                string rightWidth = "revert";
                string topWidth = "revert";
                string bottomWidth = "revert";

                string leftStyle = "revert";
                string rightStyle = "revert";
                string topStyle = "revert";
                string bottomStyle = "revert";

                string leftColor = "revert";
                string rightColor = "revert";
                string topColor = "revert";
                string bottomColor = "revert";

                if (border.LeftBorder != null)
                {
                    BorderStyleToHtmlAttributes(workbook, border.LeftBorder, ref leftWidth, ref leftStyle, ref leftColor);
                }
                if (border.RightBorder != null)
                {
                    BorderStyleToHtmlAttributes(workbook, border.RightBorder, ref rightWidth, ref rightStyle, ref rightColor);
                }
                if (border.TopBorder != null)
                {
                    BorderStyleToHtmlAttributes(workbook, border.TopBorder, ref topWidth, ref topStyle, ref topColor);
                }
                if (border.BottomBorder != null)
                {
                    BorderStyleToHtmlAttributes(workbook, border.BottomBorder, ref bottomWidth, ref bottomStyle, ref bottomColor);
                }

                styleHtml += $" border-width: {topWidth} {rightWidth} {bottomWidth} {leftWidth}; border-style: {topStyle} {rightStyle} {bottomStyle} {leftStyle}; border-color: {topColor} {rightColor} {bottomColor} {leftColor};";
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
                        case HorizontalAlignmentValues.General:
                            horizontalTextAlignment = cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == CellValues.Number ? "right" : "left";
                            break;
                    }
                }

                styleHtml += $" text-align: {horizontalTextAlignment}; vertical-align: {verticalTextAlignment};";

                if (alignment.WrapText != null && alignment.WrapText.HasValue && alignment.WrapText.Value)
                {
                    styleHtml += " word-wrap: break-word; white-space: normal;";
                }
                if (alignment.TextRotation != null && alignment.TextRotation.HasValue)
                {
                    cellValue = $"<div style=\"width: fit-content; transform: rotate(-{alignment.TextRotation.Value}deg);\">" + cellValue + "</div>";
                }
            }

            return styleHtml;
        }

        private static void BorderStyleToHtmlAttributes(WorkbookPart workbook, BorderPropertiesType border, ref string width, ref string style, ref string color)
        {
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
                string borderColor = GetColorFromColorType(workbook, border.Color);
                if (!string.IsNullOrEmpty(borderColor))
                {
                    color = borderColor;
                }
            }
        }

        private static void ConvertDrawings(WorksheetPart worksheet, OpenXmlCompositeElement anchor, StreamWriter writer, double left, double top, double width, double height, bool convertPicture)
        {
            List<DrawingInfo> drawings = new List<DrawingInfo>();

            if (convertPicture)
            {
                foreach (DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>())
                {
                    if (picture == null || picture.BlipFill == null || picture.BlipFill.Blip == null)
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
                            Prefix = $"<img src=\"data:{imagePart.ContentType};base64,{base64}\"{(properties != null && properties.Description != null && properties.Description.HasValue ? " alt=\"" + properties.Description.Value + "\"" : "")}",
                            Postfix = "/>",
                            Hidden = properties != null && properties.Hidden != null && properties.Hidden.HasValue && properties.Hidden.Value,
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
                if (shape == null)
                {
                    continue;
                }

                string text = shape.TextBody != null ? shape.TextBody.InnerText : "";

                DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties properties = null;
                if (shape.NonVisualShapeProperties != null)
                {
                    properties = shape.NonVisualShapeProperties.NonVisualDrawingProperties;
                }

                drawings.Add(new DrawingInfo()
                {
                    Prefix = "<p",
                    Postfix = $">{text}</p>",
                    Hidden = properties != null && properties.Hidden != null && properties.Hidden.HasValue && properties.Hidden.Value,
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height,
                    ShapeProperties = shape.ShapeProperties
                });
            }

            foreach (DrawingInfo drawingInfo in drawings)
            {
                double newLeft = drawingInfo.Left;
                double newTop = drawingInfo.Top;
                double newWidth = drawingInfo.Width;
                double newHeight = drawingInfo.Height;
                double rotation = 0;

                if (drawingInfo.ShapeProperties != null && drawingInfo.ShapeProperties.Transform2D != null)
                {
                    if (drawingInfo.ShapeProperties.Transform2D.Offset != null)
                    {
                        if (!double.IsNaN(newLeft))
                        {
                            newLeft += drawingInfo.ShapeProperties.Transform2D.Offset.X != null && drawingInfo.ShapeProperties.Transform2D.Offset.X.HasValue ? drawingInfo.ShapeProperties.Transform2D.Offset.X.Value / 914400 * 96 : 0;
                        }
                        if (!double.IsNaN(newTop))
                        {
                            newTop += drawingInfo.ShapeProperties.Transform2D.Offset.Y != null && drawingInfo.ShapeProperties.Transform2D.Offset.Y.HasValue ? drawingInfo.ShapeProperties.Transform2D.Offset.Y.Value / 914400 * 96 : 0;
                        }
                    }
                    if (drawingInfo.ShapeProperties.Transform2D.Extents != null)
                    {
                        if (!double.IsNaN(newWidth))
                        {
                            newWidth += drawingInfo.ShapeProperties.Transform2D.Extents.Cx != null && drawingInfo.ShapeProperties.Transform2D.Extents.Cx.HasValue ? drawingInfo.ShapeProperties.Transform2D.Extents.Cx.Value / 914400 * 96 : 0;
                        }
                        if (!double.IsNaN(newHeight))
                        {
                            newHeight += drawingInfo.ShapeProperties.Transform2D.Extents.Cy != null && drawingInfo.ShapeProperties.Transform2D.Extents.Cy.HasValue ? drawingInfo.ShapeProperties.Transform2D.Extents.Cy.Value / 914400 * 96 : 0;
                        }
                    }
                    if (drawingInfo.ShapeProperties.Transform2D.Rotation != null && drawingInfo.ShapeProperties.Transform2D.Rotation.HasValue)
                    {
                        rotation = drawingInfo.ShapeProperties.Transform2D.Rotation.Value;
                    }
                }

                writer.Write($"\n{new string(' ', 8)}{drawingInfo.Prefix} style=\"position: absolute; left: {(!double.IsNaN(newLeft) ? newLeft : 0)}px; top: {(!double.IsNaN(newTop) ? newTop : 0)}px; width: {(!double.IsNaN(newWidth) ? newWidth + "px" : "auto")}; height: {(!double.IsNaN(newHeight) ? newHeight + "px" : "auto")}px;{(rotation != 0 ? $" transform: rotate(-{rotation}deg);" : "")}{(drawingInfo.Hidden ? " visibility: hidden;" : "")}\"{drawingInfo.Postfix}");
            }
        }

        #endregion

        #region Private Structure

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
            public bool Hidden { get; set; }
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties ShapeProperties { get; set; }
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
