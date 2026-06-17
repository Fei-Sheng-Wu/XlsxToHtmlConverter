using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxToHtmlConverter
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Converter"/> class.
    /// </summary>
    public class Converter()
    {
        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The path of the local input XLSX document.</param>
        /// <param name="output">The output path of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback event handler.</param>
        public static void Convert(string input, string output, ConverterConfiguration? configuration = null, ConverterProgressChangedEventHandler? callback = null)
        {
            configuration ??= new();

            using FileStream stream = new(output, FileMode.Create, FileAccess.Write, FileShare.Read, configuration.BufferSize);
            Convert(input, stream, configuration, callback);
        }

        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The path of the local input XLSX document.</param>
        /// <param name="output">The output stream of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback event handler.</param>
        public static void Convert(string input, Stream output, ConverterConfiguration? configuration = null, ConverterProgressChangedEventHandler? callback = null)
        {
            configuration ??= new();

            using FileStream stream = new(input, FileMode.Open, FileAccess.Read, FileShare.Read, configuration.BufferSize);
            Convert(stream, output, configuration, callback);
        }

        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The stream of the input XLSX document.</param>
        /// <param name="output">The output stream of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback event handler.</param>
        public static void Convert(Stream input, Stream output, ConverterConfiguration? configuration = null, ConverterProgressChangedEventHandler? callback = null)
        {
            configuration ??= new();

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input, false);
            Convert(spreadsheet, output, configuration, callback);
        }

        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The input XLSX document.</param>
        /// <param name="output">The output stream of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback event handler.</param>
        public static void Convert(SpreadsheetDocument input, Stream output, ConverterConfiguration? configuration = null, ConverterProgressChangedEventHandler? callback = null)
        {
            configuration ??= new();
            Base.ConverterContext context = new();

            using StreamWriter writer = new(output, configuration.Encoding, configuration.BufferSize, true);
            int indent = 0;

            WorkbookPart? workbook = input.WorkbookPart;
            context.Theme = workbook?.ThemePart?.Theme;
            context.Stylesheet = configuration.ConverterComposition.XlsxStylesheetReader.Convert(workbook?.WorkbookStylesPart?.Stylesheet, context, configuration);
            context.SharedStrings = configuration.ConverterComposition.XlsxSharedStringTableReader.Convert(workbook?.SharedStringTablePart?.SharedStringTable, context, configuration);

            if (!configuration.UseHtmlFragment)
            {
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Declaration, Base.Implementation.Common.TAG_HTML), context, configuration));
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_HTML), context, configuration));
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_HEAD), context, configuration));
                indent++;

                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Unpaired, "meta", new()
                {
                    ["charset"] = configuration.Encoding.WebName
                }), context, configuration));

                if (configuration.HtmlTitle != null)
                {
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Paired, "title", null, [configuration.HtmlTitle]), context, configuration));
                }
            }

            if (configuration.ConvertStyles)
            {
                string specifier(string selector)
                {
                    return configuration.HtmlRootClass != null ? Base.Implementation.Common.Format(selector != Base.Implementation.Common.TAG_TABLE ? ".{0} {1}" : "{1}.{0}", [configuration.HtmlRootClass, selector]) : selector;
                }

                Base.Specification.Html.HtmlStylesCollection stylesheet = [];
                foreach ((string original, Base.Specification.Html.HtmlStyles styles) in configuration.HtmlPresetStylesheet)
                {
                    stylesheet[specifier(original)] = styles;
                }
                if (configuration.UseHtmlClasses)
                {
                    foreach (Base.Specification.Xlsx.XlsxBaseStyles styles in context.Stylesheet.BaseStyles)
                    {
                        stylesheet[specifier(string.Concat(".", styles.Name))] = styles.GetsStyles();
                    }
                    foreach (Base.Specification.Xlsx.XlsxDifferentialStyles styles in context.Stylesheet.DifferentialStyles)
                    {
                        stylesheet[specifier(string.Concat(".", styles.Name))] = styles.GetsStyles();
                    }
                }

                if (stylesheet.Any())
                {
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Paired, Base.Implementation.Common.TAG_STYLE, null, [stylesheet]), context, configuration));
                }
            }

            if (!configuration.UseHtmlFragment)
            {
                indent--;
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_HEAD), context, configuration));
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_BODY), context, configuration));
            }
            indent++;

            IEnumerable<Sheet> sheets = workbook?.Workbook.Sheets?.Elements<Sheet>() ?? [];
            if (configuration.XlsxSheetSelector != null)
            {
                sheets = sheets.Where((x, i) => configuration.XlsxSheetSelector((i, x.Id?.Value)));
            }

            (uint Current, uint Total) index = (1, (uint)sheets.Count());
            foreach (Sheet sheet in sheets)
            {
                WorksheetPart? worksheet = sheet.Id?.Value != null && (workbook?.TryGetPartById(sheet.Id.Value, out OpenXmlPart? part) ?? false) ? part as WorksheetPart : null;
                context.Sheet = configuration.ConverterComposition.XlsxWorksheetReader.Convert(worksheet?.Worksheet, context, configuration);

                callback?.Invoke(input, new(index, (0, context.Sheet.Dimension.RowCount)));

                context.Sheet.Specialties.AddRange(worksheet?.TableDefinitionParts.SelectMany(x => configuration.ConverterComposition.XlsxTableReader.Convert(x, context, configuration)) ?? []);
                context.Sheet.Specialties.AddRange(configuration.ConverterComposition.XlsxDrawingReader.Convert(worksheet?.DrawingsPart, context, configuration));

                HashSet<uint> lefts = [];
                HashSet<uint> tops = [];
                List<Base.Specification.Xlsx.XlsxSpecialty> elements = [];
                Dictionary<uint, List<Base.Specification.Xlsx.XlsxSpecialty>> references = [];
                foreach (Base.Specification.Xlsx.XlsxSpecialty specialty in context.Sheet.Specialties)
                {
                    if (specialty.Specialty is Base.Specification.Html.HtmlElement)
                    {
                        lefts.Add(specialty.Range.ColumnStart);
                        tops.Add(specialty.Range.RowStart);
                        lefts.Add(specialty.Range.ColumnEnd);
                        tops.Add(specialty.Range.RowEnd);
                        elements.Add(specialty);
                    }

                    for (uint i = specialty.Range.RowStart; i <= specialty.Range.RowEnd; i++)
                    {
                        if (Base.Implementation.Common.Get(references, i) is not List<Base.Specification.Xlsx.XlsxSpecialty> local)
                        {
                            local = [];
                            references[i] = local;
                        }

                        local.Add(specialty);
                    }
                }

                Base.Specification.Html.HtmlAttributes table = [];
                if (configuration.HtmlRootClass != null)
                {
                    table[Base.Implementation.Common.ATTRIBUTE_CLASS] = new Base.Specification.Html.HtmlClasses() { configuration.HtmlRootClass };
                }
                if (configuration.ConvertSizes)
                {
                    table[Base.Implementation.Common.ATTRIBUTE_STYLE] = new Base.Specification.Html.HtmlStyles(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Sheet, configuration.UseHtmlProportionalWidths ? Base.Specification.Html.HtmlStyleType.WidthFull : Base.Specification.Html.HtmlStyleType.WidthFit), context, configuration));
                }
                if (sheet.State?.Value != null && sheet.State.Value != SheetStateValues.Visible && configuration.ConvertVisibilities)
                {
                    table[Base.Implementation.Common.ATTRIBUTE_HIDDEN] = null;
                }

                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_TABLE, table), context, configuration));
                indent++;

                if (configuration.ConvertSheetTitles)
                {
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Paired, Base.Implementation.Common.TAG_CAPTION, context.Sheet.TitleAttributes, sheet.Name?.Value != null ? [sheet.Name.Value] : null), context, configuration));
                }

                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_COLUMN_GROUP), context, configuration));
                indent++;

                (double Width, bool? IsHidden, uint? StylesIndex)[] columns = new (double Width, bool? IsHidden, uint? StylesIndex)[context.Sheet.Dimension.ColumnCount];
                for (uint i = 0; i < columns.Length; i++)
                {
                    (double? width, bool? isHidden, uint? styles) = Base.Implementation.Common.Get(context.Sheet.Columns, i);
                    columns[i] = (width ?? context.Sheet.CellSize.Width, isHidden, styles);
                }
                if (configuration.UseHtmlProportionalWidths)
                {
                    double total = columns.Sum(x => double.IsNaN(x.Width) ? context.Sheet.CellSize.Width : x.Width);
                    columns = [.. columns.Select(x => (100.0 * x.Width / total, x.IsHidden, x.StylesIndex))];
                }

                for (uint i = 0; i < columns.Length; i++)
                {
                    Base.Specification.Html.HtmlStyles baseline = [];
                    Base.Specification.Html.HtmlAttributes attributes = new()
                    {
                        [Base.Implementation.Common.ATTRIBUTE_STYLE] = baseline
                    };
                    attributes.Merge(context.Sheet.ColumnAttributes);
                    if (configuration.ConvertSizes)
                    {
                        baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Column, (double.IsNaN(columns[i].Width), configuration.UseHtmlProportionalWidths) switch
                        {
                            (false, false) => Base.Specification.Html.HtmlStyleType.WidthNumeric,
                            (false, true) => Base.Specification.Html.HtmlStyleType.WidthProportional,
                            _ => Base.Specification.Html.HtmlStyleType.WidthAutomatic
                        }, Base.Implementation.Common.Format(columns[i].Width, configuration)), context, configuration));
                    }
                    if (columns[i].IsHidden != null && configuration.ConvertVisibilities)
                    {
                        baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Column, (columns[i].IsHidden ?? false) ? Base.Specification.Html.HtmlStyleType.VisibilityCollapsed : Base.Specification.Html.HtmlStyleType.VisibilityVisible), context, configuration));
                    }

                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Unpaired, Base.Implementation.Common.TAG_COLUMN, attributes), context, configuration));
                }

                indent--;
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_COLUMN_GROUP), context, configuration));
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_ROW_GROUP), context, configuration));
                indent++;

                (uint Column, uint Row) last = (context.Sheet.Dimension.ColumnStart - 1, context.Sheet.Dimension.RowStart - 1);
                List<Base.Specification.Xlsx.XlsxSpecialty> specialties = [];

                void populator(uint column, uint row, Base.Specification.Html.HtmlElement? element = null)
                {
                    if (specialties.Any(x => x.Specialty is MergeCell && x.Range.ContainsColumn(column) && !x.Range.StartsAt(column, row) && context.Sheet.Dimension.Contains(x.Range.ColumnStart, x.Range.RowStart)))
                    {
                        return;
                    }

                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(element ?? new(indent, Base.Specification.Html.HtmlElementType.Paired, Base.Implementation.Common.TAG_CELL), context, configuration));
                }
                void padder()
                {
                    if (last.Row < context.Sheet.Dimension.RowStart)
                    {
                        return;
                    }

                    for (uint i = last.Column + 1; i <= context.Sheet.Dimension.ColumnEnd; i++)
                    {
                        populator(i, last.Row);
                    }

                    indent--;
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_ROW), context, configuration));

                    callback?.Invoke(input, new(index, (last.Row - context.Sheet.Dimension.RowStart + 1, context.Sheet.Dimension.RowCount)));
                }

                Row? row = null;
                foreach (Base.Specification.Xlsx.XlsxCell entry in configuration.ConverterComposition.XlsxWorksheetIterator.Convert(context.Sheet, context, configuration))
                {
                    Base.Specification.Xlsx.XlsxCell cell = entry;
                    if (!context.Sheet.Dimension.Contains(cell.Reference.Column, cell.Reference.Row))
                    {
                        continue;
                    }

                    while (cell.Reference.Row > last.Row)
                    {
                        padder();
                        last = (context.Sheet.Dimension.ColumnStart - 1, last.Row + 1);
                        specialties = Base.Implementation.Common.Get(references, last.Row) ?? [];
                        row = cell.Reference.Row <= last.Row ? cell.Cell?.Parent as Row : null;

                        Base.Specification.Html.HtmlStyles baseline = [];
                        Base.Specification.Html.HtmlAttributes attributes = new()
                        {
                            [Base.Implementation.Common.ATTRIBUTE_STYLE] = baseline
                        };
                        attributes.Merge(context.Sheet.RowAttributes);
                        if (configuration.ConvertSizes)
                        {
                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Row, Base.Specification.Html.HtmlStyleType.HeightNumeric, Base.Implementation.Common.Format((Base.Implementation.Common.Get(row?.Height?.Value, row != null ? row?.CustomHeight?.Value : false) * Base.Implementation.Common.RATIO_POINT) ?? context.Sheet.CellSize.Height, configuration)), context, configuration));
                        }
                        if (row?.Hidden?.Value != null && configuration.ConvertVisibilities)
                        {
                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Row, row.Hidden.Value ? Base.Specification.Html.HtmlStyleType.VisibilityCollapsed : Base.Specification.Html.HtmlStyleType.VisibilityVisible), context, configuration));
                        }
                        if (tops.Contains(last.Row))
                        {
                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Row, Base.Specification.Html.HtmlStyleType.AnchorDefinitionNumeric, Base.Implementation.Common.Format(last.Row, configuration)), context, configuration));
                        }

                        writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_ROW, attributes), context, configuration));
                        indent++;
                    }
                    for (uint i = last.Column + 1; i < cell.Reference.Column; i++)
                    {
                        populator(i, cell.Reference.Row);
                    }

                    bool isSelected = configuration.XlsxCellSelector?.Invoke((cell.Reference.Column, cell.Reference.Row)) ?? true;

                    Base.Specification.Xlsx.XlsxBaseStyles? shared = Base.Implementation.Common.Get(context.Stylesheet.BaseStyles, cell.Cell?.StyleIndex?.Value ?? Base.Implementation.Common.Get(columns, cell.Reference.Column - context.Sheet.Dimension.ColumnStart).StylesIndex ?? row?.StyleIndex?.Value ?? 0);
                    if (shared != null)
                    {
                        cell.Styles.Add(shared);
                        cell.NumberFormat = shared.NumberFormatIdentifier != null ? Base.Implementation.Common.Get(context.Stylesheet.NumberFormats, shared.NumberFormatIdentifier.Value) : null;
                        cell.NumberFormatIdentifier = shared.NumberFormatIdentifier;
                    }

                    Base.Specification.Html.HtmlStyles individual = [];
                    cell.Specialties = specialties.Where(x => x.Range.ContainsColumn(cell.Reference.Column));
                    foreach (Base.Specification.Xlsx.XlsxSpecialty specialty in cell.Specialties)
                    {
                        switch (specialty.Specialty)
                        {
                            case MergeCell when specialty.Range.StartsAt(cell.Reference.Column, cell.Reference.Row):
                                cell.Attributes["colspan"] = Base.Implementation.Common.Format(specialty.Range.ColumnCount, configuration);
                                cell.Attributes["rowspan"] = Base.Implementation.Common.Format(specialty.Range.RowCount, configuration);
                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Cell, Base.Specification.Html.HtmlStyleType.ClippingHorizontalHidden), context, configuration));
                                break;
                            case Base.Specification.Xlsx.XlsxStyles styles when isSelected:
                                cell.Styles.Add(styles);
                                if (styles is Base.Specification.Xlsx.XlsxDifferentialStyles differential && differential.NumberFormat != null && configuration.ConvertNumberFormats)
                                {
                                    cell.NumberFormat = differential.NumberFormat;
                                }
                                break;
                        }
                    }

                    if (configuration.UseHtmlClasses)
                    {
                        cell.Attributes[Base.Implementation.Common.ATTRIBUTE_CLASS] = new Base.Specification.Html.HtmlClasses();
                    }
                    cell.Attributes[Base.Implementation.Common.ATTRIBUTE_STYLE] = individual;
                    cell.Attributes.Merge(context.Sheet.CellAttributes);

                    if (isSelected)
                    {
                        cell = configuration.ConverterComposition.XlsxCellContentReader.Convert(cell, context, configuration);
                    }
                    Base.Specification.Html.HtmlElement element = new(indent, Base.Specification.Html.HtmlElementType.Paired, Base.Implementation.Common.TAG_CELL, cell.Attributes, cell.Children);

                    bool isHidden = false;
                    foreach (Base.Specification.Xlsx.XlsxStyles styles in isSelected ? cell.Styles : [])
                    {
                        if (configuration.ConvertStyles)
                        {
                            styles.ApplyStyles(element, configuration.UseHtmlClasses);
                        }

                        if (styles.IsHidden != null && configuration.ConvertVisibilities)
                        {
                            isHidden = styles.IsHidden.Value;
                        }
                    }
                    if (isHidden && configuration.ConvertVisibilities)
                    {
                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Cell, Base.Specification.Html.HtmlStyleType.VisibilityHidden), context, configuration));
                    }

                    populator(cell.Reference.Column, cell.Reference.Row, element);
                    last = cell.Reference;
                }
                padder();

                if (elements.Any())
                {
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_ROW, new()
                    {
                        [Base.Implementation.Common.ATTRIBUTE_STYLE] = new Base.Specification.Html.HtmlStyles(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Row, Base.Specification.Html.HtmlStyleType.VisibilityCollapsed), context, configuration))
                    }), context, configuration));
                    indent++;

                    Base.Specification.Html.HtmlAttributes? anchor(uint column)
                    {
                        Base.Specification.Html.HtmlStyles styles = [];
                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Base.Specification.Html.HtmlStyleTarget.Column, Base.Specification.Html.HtmlStyleType.AnchorDefinitionNumeric, Base.Implementation.Common.Format(column, configuration)), context, configuration));

                        return lefts.Contains(column) ? new()
                        {
                            [Base.Implementation.Common.ATTRIBUTE_STYLE] = styles
                        } : null;
                    }

                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedStart, Base.Implementation.Common.TAG_CELL, anchor(context.Sheet.Dimension.ColumnStart)), context, configuration));
                    indent++;

                    foreach (Base.Specification.Xlsx.XlsxSpecialty specialty in elements)
                    {
                        if (specialty.Specialty is not Base.Specification.Html.HtmlElement element)
                        {
                            continue;
                        }

                        Base.Specification.Html.HtmlStyles positions = [];
                        if (specialty.Range.RowStart > 0)
                        {
                            positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Base.Specification.Html.HtmlStyleType.VariablePositionTopAnchoring, Base.Implementation.Common.Format(specialty.Range.RowStart, configuration)), context, configuration));
                        }
                        if (specialty.Range.ColumnEnd > 0)
                        {
                            positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Base.Specification.Html.HtmlStyleType.VariablePositionRightAnchoring, Base.Implementation.Common.Format(specialty.Range.ColumnEnd, configuration)), context, configuration));
                        }
                        if (specialty.Range.RowEnd > 0)
                        {
                            positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Base.Specification.Html.HtmlStyleType.VariablePositionBottomAnchoring, Base.Implementation.Common.Format(specialty.Range.RowEnd, configuration)), context, configuration));
                        }
                        if (specialty.Range.ColumnStart > 0)
                        {
                            positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Base.Specification.Html.HtmlStyleType.VariablePositionLeftAnchoring, Base.Implementation.Common.Format(specialty.Range.ColumnStart, configuration)), context, configuration));
                        }
                        positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Base.Specification.Html.HtmlStyleType.VisibilityVisible), context, configuration));

                        element.Indent = indent;
                        element.Attributes.MergeStyles(positions);

                        writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(element, context, configuration));
                    }

                    indent--;
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_CELL), context, configuration));

                    for (uint i = context.Sheet.Dimension.ColumnStart + 1; i <= context.Sheet.Dimension.ColumnEnd; i++)
                    {
                        writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.Paired, Base.Implementation.Common.TAG_CELL, anchor(i)), context, configuration));
                    }

                    indent--;
                    writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_ROW), context, configuration));
                }

                indent--;
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_ROW_GROUP), context, configuration));
                indent--;
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_TABLE), context, configuration));

                index = (index.Current + 1, index.Total);
            }

            indent--;
            if (!configuration.UseHtmlFragment)
            {
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_BODY), context, configuration));
                writer.Write(configuration.ConverterComposition.HtmlWriter.Convert(new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, Base.Implementation.Common.TAG_HTML), context, configuration));
            }
        }
    }
}

namespace XlsxToHtmlConverter.Base.Implementation
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Common"/> class.
    /// </summary>
    public class Common()
    {
        /// <summary>
        /// Represents the angle ratio.
        /// </summary>
        public const double RATIO_ANGLE = 1 / 60000.0;

        /// <summary>
        /// Represents the percentage ratio.
        /// </summary>
        public const double RATIO_PERCENTAGE = 1 / 100000.0;

        /// <summary>
        /// Represents the point ratio.
        /// </summary>
        public const double RATIO_POINT = 1 / 72.0 * 96.0;

        /// <summary>
        /// Represents the point spacing ratio.
        /// </summary>
        public const double RATIO_POINT_SPACING = 1 / 7200.0 * 96.0;

        /// <summary>
        /// Represents the English Metric Unit ratio.
        /// </summary>
        public const double RATIO_ENGLISH_METRIC_UNIT = 1 / 914400.0 * 96.0;

        /// <summary>
        /// Represents the document tag.
        /// </summary>
        public const string TAG_HTML = "html";

        /// <summary>
        /// Represents the document head tag.
        /// </summary>
        public const string TAG_HEAD = "head";

        /// <summary>
        /// Represents the document body tag.
        /// </summary>
        public const string TAG_BODY = "body";

        /// <summary>
        /// Represents the style tag.
        /// </summary>
        public const string TAG_STYLE = "style";

        /// <summary>
        /// Represents the table tag.
        /// </summary>
        public const string TAG_TABLE = "table";

        /// <summary>
        /// Represents the caption tag.
        /// </summary>
        public const string TAG_CAPTION = "caption";

        /// <summary>
        /// Represents the column tag.
        /// </summary>
        public const string TAG_COLUMN = "col";

        /// <summary>
        /// Represents the column group tag.
        /// </summary>
        public const string TAG_COLUMN_GROUP = "colgroup";

        /// <summary>
        /// Represents the row tag.
        /// </summary>
        public const string TAG_ROW = "tr";

        /// <summary>
        /// Represents the row group tag.
        /// </summary>
        public const string TAG_ROW_GROUP = "tbody";

        /// <summary>
        /// Represents the cell tag.
        /// </summary>
        public const string TAG_CELL = "td";

        /// <summary>
        /// Represents the container tag.
        /// </summary>
        public const string TAG_CONTAINER = "div";

        /// <summary>
        /// Represents the text tag.
        /// </summary>
        public const string TAG_TEXT = "span";

        /// <summary>
        /// Represents the text group tag.
        /// </summary>
        public const string TAG_TEXT_GROUP = "p";

        /// <summary>
        /// Represents the style attribute.
        /// </summary>
        public const string ATTRIBUTE_STYLE = "style";

        /// <summary>
        /// Represents the class attribute.
        /// </summary>
        public const string ATTRIBUTE_CLASS = "class";

        /// <summary>
        /// Represents the hidden attribute.
        /// </summary>
        public const string ATTRIBUTE_HIDDEN = "hidden";

        /// <summary>
        /// Specifies the category of a cached entry.
        /// </summary>
        public enum CacheCategory
        {
            /// <summary>
            /// Common styles.
            /// </summary>
            CommonStyles,

            /// <summary>
            /// Number format.
            /// </summary>
            NumberFormat
        }

        /// <summary>
        /// Retrieves a specified value.
        /// </summary>
        /// <typeparam name="T">The type of the value.</typeparam>
        /// <param name="value">The specified value.</param>
        /// <param name="flag">Whether the value can be retrieved.</param>
        /// <returns>The retrieved value.</returns>
        public static T? Get<T>(T? value, bool? flag)
        {
            return (flag ?? true) ? value : default;
        }

        /// <summary>
        /// Retrieves the value at a specified index within an <see cref="Array"/>.
        /// </summary>
        /// <typeparam name="T">The type of the value.</typeparam>
        /// <param name="values">The <see cref="Array"/> instance to retrieve from.</param>
        /// <param name="index">The specified index.</param>
        /// <param name="flag">Whether the value can be retrieved.</param>
        /// <returns>The retrieved value.</returns>
        public static T? Get<T>(T?[] values, uint? index, bool? flag = null)
        {
            return index != null && index.Value < values.Length ? Get(values[index.Value], flag) : default;
        }

        /// <summary>
        /// Retrieves the value with a specified key within a <see cref="Dictionary{T1, T2}"/>.
        /// </summary>
        /// <typeparam name="T1">The type of the key.</typeparam>
        /// <typeparam name="T2">The type of the value.</typeparam>
        /// <param name="values">The <see cref="Dictionary{T1, T2}"/> instance to retrieve from.</param>
        /// <param name="key">The specified key.</param>
        /// <param name="flag">Whether the value can be retrieved.</param>
        /// <returns>The retrieved value.</returns>
        public static T2? Get<T1, T2>(Dictionary<T1, T2> values, T1? key, bool? flag = null) where T1 : notnull
        {
            return key != null && values.TryGetValue(key, out T2? result) ? Get(result, flag) : default;
        }

        /// <summary>
        /// Formats a numeric value.
        /// </summary>
        /// <param name="value">The numeric value.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <returns>The formatted result.</returns>
        public static string Format(object value, ConverterConfiguration configuration)
        {
            return value switch
            {
                uint integer => integer.ToString(CultureInfo.InvariantCulture),
                int integer => integer.ToString(CultureInfo.InvariantCulture),
                long integer => integer.ToString(CultureInfo.InvariantCulture),
                double decimals => decimals.ToString(configuration.RoundingDigits < 0 ? "G" : string.Concat("0.", new string('#', configuration.RoundingDigits)), CultureInfo.InvariantCulture),
                _ => value.ToString() ?? string.Empty
            };
        }

        /// <summary>
        /// Formats a collection of values using the specified format.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <param name="values">The collection of values.</param>
        /// <returns>The formatted result.</returns>
        public static string Format(string format, string[] values)
        {
            return string.Format(CultureInfo.InvariantCulture, format, values);
        }

        /// <summary>
        /// Converts a string representation to a numeric value.
        /// </summary>
        /// <param name="value">The string representation.</param>
        /// <returns>The numeric value.</returns>
        public static uint? ParsePositive(string? value)
        {
            return uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint result) ? result : null;
        }

        /// <summary>
        /// Converts a string representation to a numeric value.
        /// </summary>
        /// <param name="value">The string representation.</param>
        /// <returns>The numeric value.</returns>
        public static int? ParseInteger(string? value)
        {
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result) ? result : null;
        }

        /// <summary>
        /// Converts a string representation to a numeric value.
        /// </summary>
        /// <param name="value">The string representation.</param>
        /// <returns>The numeric value.</returns>
        public static long? ParseLarge(string? value)
        {
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long result) ? result : null;
        }

        /// <summary>
        /// Converts a string representation to a numeric value.
        /// </summary>
        /// <param name="value">The string representation.</param>
        /// <returns>The numeric value.</returns>
        public static int? ParseHex(string? value)
        {
            return int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int result) ? result : null;
        }

        /// <summary>
        /// Converts a string representation to a numeric value.
        /// </summary>
        /// <param name="value">The string representation.</param>
        /// <returns>The numeric value.</returns>
        public static double? ParseDecimals(string? value)
        {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : null;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultHtmlWriter"/> class.
    /// </summary>
    public class DefaultHtmlWriter() : IConverterBase<Specification.Html.HtmlElement, string>
    {
        /// <inheritdoc />
        public string Convert(Specification.Html.HtmlElement value, ConverterContext context, ConverterConfiguration configuration)
        {
            string padder(int indent)
            {
                return string.Concat(Enumerable.Repeat(configuration.TabCharacter, indent));
            }
            string compositor(Specification.Html.HtmlElement element)
            {
                string attributes = string.Concat(element.Attributes.Select(x => x.Value switch
                {
                    string value => Common.Format(" {0}=\"{1}\"", [x.Key, value]),
                    Specification.Html.HtmlClasses classes => classes.Any() ? Common.Format(" {0}=\"{1}\"", [x.Key, string.Join(' ', classes)]) : string.Empty,
                    Specification.Html.HtmlStyles styles => styles.Any() ? Common.Format(" {0}=\"{1}\"", [x.Key, string.Join(' ', styles.Select(y => Common.Format("{0}: {1};", [y.Key, y.Value])))]) : string.Empty,
                    _ => string.Concat(" ", x.Key),
                }));

                return element.Type switch
                {
                    Specification.Html.HtmlElementType.Declaration => Common.Format("<!DOCTYPE {0}>", [element.Tag]),
                    Specification.Html.HtmlElementType.Paired => Common.Format("<{0}{1}>{2}</{0}>", [element.Tag, attributes, populator(element.Children, element.Indent ?? 0)]),
                    Specification.Html.HtmlElementType.PairedStart => Common.Format("<{0}{1}>", [element.Tag, attributes]),
                    Specification.Html.HtmlElementType.PairedEnd => Common.Format("</{0}>", [element.Tag]),
                    Specification.Html.HtmlElementType.Unpaired => Common.Format("<{0}{1}>", [element.Tag, attributes]),
                    _ => Common.Format("<!-- {0} -->", [populator(element.Children, element.Indent ?? 0)])
                };
            }
            string populator(Specification.Html.HtmlChildren content, int indent)
            {
                return string.Concat(content.Select(x =>
                {
                    switch (x)
                    {
                        case Specification.Html.HtmlElement html:
                            return compositor(html);
                        case Specification.Html.HtmlStylesCollection css:
                            StringBuilder builder = new(configuration.NewlineCharacter);
                            foreach ((string selector, Specification.Html.HtmlStyles styles) in css)
                            {
                                builder.Append(string.Concat(padder(indent + 1), selector, " {", configuration.NewlineCharacter));
                                foreach ((string property, string value) in styles)
                                {
                                    builder.Append(string.Concat(padder(indent + 2), Common.Format("{0}: {1};", [property, value]), configuration.NewlineCharacter));
                                }
                                builder.Append(string.Concat(padder(indent + 1), "}", configuration.NewlineCharacter));
                            }
                            builder.Append(padder(indent));

                            return builder.ToString();
                        default:
                            return WebUtility.HtmlEncode(x.ToString()) ?? string.Empty;
                    }
                }));
            }

            return string.Concat(padder(value.Indent ?? 0), compositor(value), configuration.NewlineCharacter);
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultHtmlStylizer"/> class.
    /// </summary>
    public class DefaultHtmlStylizer() : IConverterBase<Specification.Html.HtmlStyleDefinition, IEnumerable<KeyValuePair<string, string>>>
    {
        /// <inheritdoc />
        public IEnumerable<KeyValuePair<string, string>> Convert(Specification.Html.HtmlStyleDefinition value, ConverterContext context, ConverterConfiguration configuration)
        {
            return (value.Target, value.Type) switch
            {
                (Specification.Html.HtmlStyleTarget.SheetTitle, Specification.Html.HtmlStyleType.VariableSheetColorExact) => [new("--sheet-color", value.Parameter ?? string.Empty)],
                (Specification.Html.HtmlStyleTarget.Column, Specification.Html.HtmlStyleType.WidthNumeric) => [new("width", string.Concat(value.Parameter ?? string.Empty, "ch"))],
                (Specification.Html.HtmlStyleTarget.Column, Specification.Html.HtmlStyleType.WidthProportional) => [new("width", string.Concat(value.Parameter ?? string.Empty, "%"))],
                (Specification.Html.HtmlStyleTarget.Column, Specification.Html.HtmlStyleType.AnchorDefinitionNumeric) => [new("anchor-name", string.Concat("--column-", value.Parameter ?? string.Empty))],
                (Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.HeightNumeric) => [new("height", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.BorderThicknessTopThick) => [new("border-top-width", "thick")],
                (Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.BorderThicknessBottomThick) => [new("border-bottom-width", "thick")],
                (Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.AnchorDefinitionNumeric) => [new("anchor-name", string.Concat("--row-", value.Parameter ?? string.Empty))],
                (Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.TextIndentNumeric) => [new("padding-inline-start", string.Concat(value.Parameter ?? string.Empty, "ch"))],
                (Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.AlignmentHorizontalDistributed) => [new("display", "flex"), new("justify-content", "space-between")],
                (Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ClippingHorizontalHidden) => [new("overflow-x", "hidden")],
                (Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.AlignmentVerticalTop) => [new("align-content", "start")],
                (Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.AlignmentVerticalCenter) => [new("align-content", "center")],
                (Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.AlignmentVerticalBottom) => [new("align-content", "end")],
                (Specification.Html.HtmlStyleTarget.Cell or Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.RotationNumeric) => [new("display", "inline-block"), new("rotate", string.Concat(value.Parameter ?? string.Empty, "deg"))],
                (_, Specification.Html.HtmlStyleType.VariablePositionTopAnchoring) => [new("--top", Common.Format("anchor(--row-{0} top)", [value.Parameter ?? string.Empty]))],
                (_, Specification.Html.HtmlStyleType.VariablePositionRightAnchoring) => [new("--right", Common.Format("anchor(--column-{0} left)", [value.Parameter ?? string.Empty]))],
                (_, Specification.Html.HtmlStyleType.VariablePositionBottomAnchoring) => [new("--bottom", Common.Format("anchor(--row-{0} top)", [value.Parameter ?? string.Empty]))],
                (_, Specification.Html.HtmlStyleType.VariablePositionLeftAnchoring) => [new("--left", Common.Format("anchor(--column-{0} left)", [value.Parameter ?? string.Empty]))],
                (_, Specification.Html.HtmlStyleType.PositioningAbsolute) => [new("position", "absolute")],
                (_, Specification.Html.HtmlStyleType.MarginAllExact) => [new("margin", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.MarginTopExact) => [new("margin-top", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.MarginRightExact) => [new("margin-right", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.MarginBottomExact) => [new("margin-bottom", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.MarginLeftExact) => [new("margin-left", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.PaddingAllExact) => [new("padding", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.PaddingTopExact) => [new("padding-top", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.PaddingRightExact) => [new("padding-right", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.PaddingBottomExact) => [new("padding-bottom", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.PaddingLeftExact) => [new("padding-left", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.WidthFit) => [new("width", "fit-content")],
                (_, Specification.Html.HtmlStyleType.WidthAutomatic) => [new("width", "auto")],
                (_, Specification.Html.HtmlStyleType.WidthFull) => [new("width", "100%")],
                (_, Specification.Html.HtmlStyleType.HeightFull) => [new("height", "100%")],
                (_, Specification.Html.HtmlStyleType.TranslationTopExact) => [new("top", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TranslationRightExact) => [new("right", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TranslationBottomExact) => [new("bottom", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TranslationLeftExact) => [new("left", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.RotationNumeric) => [new("rotate", string.Concat(value.Parameter ?? string.Empty, "deg"))],
                (_, Specification.Html.HtmlStyleType.ScalingExact) => [new("scale", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TextSizeNumeric) => [new("font-size", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.TextFamilyTextual) => [new("font-family", string.Concat("\'", value.Parameter, "\'"))],
                (_, Specification.Html.HtmlStyleType.TextWeightNormal) => [new("font-weight", "normal")],
                (_, Specification.Html.HtmlStyleType.TextWeightBold) => [new("font-weight", "bold")],
                (_, Specification.Html.HtmlStyleType.TextStyleNormal) => [new("font-style", "normal")],
                (_, Specification.Html.HtmlStyleType.TextStyleItalic) => [new("font-style", "italic")],
                (_, Specification.Html.HtmlStyleType.TextStretchNormal) => [new("font-stretch", "normal")],
                (_, Specification.Html.HtmlStyleType.TextStretchExpanded) => [new("font-stretch", "expanded")],
                (_, Specification.Html.HtmlStyleType.TextStretchCondensed) => [new("font-stretch", "condensed")],
                (_, Specification.Html.HtmlStyleType.TextDecorationExact) => [new("text-decoration", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TextTransformExact) => [new("text-transform", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TextLetterSpacingNumeric) => [new("letter-spacing", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.TextLineHeightNumeric) => [new("line-height", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.TextLineHeightProportional) => [new("line-height", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TextIndentNumeric) => [new("text-indent", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.TextTabNumeric) => [new("tab-size", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.TextColumnCountNumeric) => [new("column-count", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.TextColumnGapNumeric) => [new("column-gap", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.TextWrappingNone) => [new("white-space", "preserve nowrap")],
                (_, Specification.Html.HtmlStyleType.TextWrappingWrap) => [new("white-space", "preserve wrap")],
                (_, Specification.Html.HtmlStyleType.TextOrientationVertical) => [new("text-orientation", "upright"), new("writing-mode", "vertical-rl")],
                (_, Specification.Html.HtmlStyleType.TextOrientationVerticalReverse) => [new("text-orientation", "upright"), new("writing-mode", "vertical-lr")],
                (_, Specification.Html.HtmlStyleType.ForegroundExact) => [new("color", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.ForegroundNone) => [new("color", "transparent")],
                (_, Specification.Html.HtmlStyleType.BackgroundExact) => [new("background", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BackgroundNone) => [new("background", "transparent")],
                (_, Specification.Html.HtmlStyleType.BorderAllRegular) => [new("border", Common.Format("thin solid {0}", [value.Parameter ?? string.Empty]))],
                (_, Specification.Html.HtmlStyleType.BorderTopExact) => [new("border-top", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BorderRightExact) => [new("border-right", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BorderBottomExact) => [new("border-bottom", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BorderLeftExact) => [new("border-left", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BorderColorAllExact) => [new("border-color", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BorderThicknessAllNumeric) => [new("border-width", string.Concat(value.Parameter ?? string.Empty, "px"))],
                (_, Specification.Html.HtmlStyleType.BorderStyleAllSolid) => [new("border-style", "solid")],
                (_, Specification.Html.HtmlStyleType.BorderStyleAllDouble) => [new("border-style", "double")],
                (_, Specification.Html.HtmlStyleType.BorderStyleAllDashed) => [new("border-style", "dashed")],
                (_, Specification.Html.HtmlStyleType.BorderStyleAllDotted) => [new("border-style", "dotted")],
                (_, Specification.Html.HtmlStyleType.AlignmentHorizontalLeft) => [new("text-align", "left")],
                (_, Specification.Html.HtmlStyleType.AlignmentHorizontalCenter) => [new("text-align", "center")],
                (_, Specification.Html.HtmlStyleType.AlignmentHorizontalRight) => [new("text-align", "right")],
                (_, Specification.Html.HtmlStyleType.AlignmentHorizontalJustify) => [new("text-align", "justify")],
                (_, Specification.Html.HtmlStyleType.AlignmentVerticalTop) => [new("vertical-align", "top")],
                (_, Specification.Html.HtmlStyleType.AlignmentVerticalCenter) => [new("vertical-align", "middle")],
                (_, Specification.Html.HtmlStyleType.AlignmentVerticalBottom) => [new("vertical-align", "bottom")],
                (_, Specification.Html.HtmlStyleType.AlignmentVerticalBaseline) => [new("vertical-align", "baseline")],
                (_, Specification.Html.HtmlStyleType.AlignmentVerticalSuperscript) => [new("vertical-align", "super")],
                (_, Specification.Html.HtmlStyleType.AlignmentVerticalSubscript) => [new("vertical-align", "sub")],
                (_, Specification.Html.HtmlStyleType.ClippingAllHidden) => [new("overflow", "clip")],
                (_, Specification.Html.HtmlStyleType.ClippingAllPath) => [new("clip-path", Common.Format("path(\'{0}\')", [value.Parameter ?? string.Empty]))],
                (_, Specification.Html.HtmlStyleType.ClippingHorizontalVisible) => [new("overflow-x", "visible")],
                (_, Specification.Html.HtmlStyleType.ClippingHorizontalHidden) => [new("overflow-x", "clip")],
                (_, Specification.Html.HtmlStyleType.ClippingVerticalVisible) => [new("overflow-y", "visible")],
                (_, Specification.Html.HtmlStyleType.ClippingVerticalHidden) => [new("overflow-y", "clip")],
                (_, Specification.Html.HtmlStyleType.CroppingExact) => [new("object-view-box", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.BoundingInclusive) => [new("box-sizing", "border-box")],
                (_, Specification.Html.HtmlStyleType.DirectionExact) => [new("direction", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.FilterExact) => [new("filter", value.Parameter ?? string.Empty)],
                (_, Specification.Html.HtmlStyleType.VisibilityVisible) => [new("visibility", "visible")],
                (_, Specification.Html.HtmlStyleType.VisibilityHidden) => [new("content-visibility", "hidden")],
                (_, Specification.Html.HtmlStyleType.VisibilityCollapsed) => [new("visibility", "collapse")],
                _ => []
            };
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxWorksheetIterator"/> class.
    /// </summary>
    public class DefaultXlsxWorksheetIterator() : IConverterBase<Specification.Xlsx.XlsxSheet?, IEnumerable<Specification.Xlsx.XlsxCell>>
    {
        /// <inheritdoc />
        public IEnumerable<Specification.Xlsx.XlsxCell> Convert(Specification.Xlsx.XlsxSheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                yield break;
            }

            (uint Column, uint Row) last = (value.Dimension.ColumnStart - 1, value.Dimension.RowStart - 1);

            (uint Column, uint Row) locater(string reference)
            {
                (uint column, uint row) = Specification.Xlsx.XlsxRange.ParseReference(reference);
                if (row < last.Row || (row == last.Row && column <= last.Column))
                {
                    column = last.Column + 1;
                    row = last.Row;
                }

                return (column, row);
            }

            foreach (Row row in value.Data?.Elements<Row>() ?? [])
            {
                uint index = row.RowIndex != null ? Math.Max(last.Row, row.RowIndex.Value) : (last.Row + 1);
                foreach (Cell cell in row.Elements<Cell>())
                {
                    last = cell.CellReference?.Value != null ? locater(cell.CellReference.Value) : (index > last.Row ? value.Dimension.ColumnStart : last.Column + 1, index);

                    yield return new(cell)
                    {
                        Reference = last
                    };
                }
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxStylesheetReader"/> class.
    /// </summary>
    public class DefaultXlsxStylesheetReader() : IConverterBase<Stylesheet?, Specification.Xlsx.XlsxStylesCollection>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxStylesCollection Convert(Stylesheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesCollection result = new();

            Specification.Xlsx.XlsxNumberFormat? parser(string? raw, uint? identifier)
            {
                if (WebUtility.HtmlDecode(raw) is not string code || code.All(char.IsWhiteSpace))
                {
                    return null;
                }

                StringBuilder builder = new();
                List<Specification.Xlsx.XlsxNumberFormatCode> codes = [new()];
                foreach ((int index, char character, bool isEscaped) in Specification.Xlsx.XlsxNumberFormat.Escape(code, null, ['[', ']']))
                {
                    if (!isEscaped)
                    {
                        switch (char.ToUpperInvariant(character))
                        {
                            case ';':
                                codes[^1]?.Code = builder.ToString();
                                codes.Add(new());
                                builder.Clear();
                                continue;
                            case 'Y' or 'M' or 'D' or 'H' or 'S':
                                codes[^1]?.IsCalendrical = true;
                                break;
                        }
                    }

                    builder.Append(character);
                }
                codes[^1]?.Code = builder.ToString();

                Specification.Xlsx.XlsxNumberFormat format = codes.Count switch
                {
                    1 => new(codes[0]),
                    2 => new(codes[0], codes[1]),
                    3 => new(codes[0], codes[1], codes[2]),
                    _ => new(codes[0], codes[1], codes[2], codes[3])
                };
                if (identifier != null)
                {
                    result.NumberFormats[identifier.Value] = format;
                }

                return format;
            }

            Font?[] fonts = [.. Common.Get(value.Fonts, configuration.ConvertStyles)?.Elements().Select(x => x as Font) ?? []];
            Fill?[] fills = [.. Common.Get(value.Fills, configuration.ConvertStyles)?.Elements().Select(x => x as Fill) ?? []];
            Border?[] borders = [.. Common.Get(value.Borders, configuration.ConvertStyles)?.Elements().Select(x => x as Border) ?? []];
            result.BaseStyles = [.. value.CellFormats?.Elements().Select((x, i) =>
            {
                if (x is not CellFormat cell)
                {
                    return new();
                }

                Specification.Xlsx.XlsxBaseStyles styles = new()
                {
                    Name = configuration.UseHtmlClasses ? string.Concat("base-", Common.Format(i, configuration)) : null,
                    IsHidden = Common.Get(cell.Protection?.Hidden?.Value, configuration.ConvertVisibilities ? cell.ApplyProtection?.Value : false)
                };
                if (configuration.ConvertStyles)
                {
                    styles.Styles.Merge(configuration.ConverterComposition.XlsxFontConverter.Convert(Common.Get(fonts, cell.FontId?.Value, cell.ApplyFont?.Value), context, configuration));
                    styles.Styles.Merge(configuration.ConverterComposition.XlsxFillConverter.Convert(Common.Get(fills, cell.FillId?.Value, cell.ApplyFill?.Value), context, configuration));
                    styles.Styles.Merge(configuration.ConverterComposition.XlsxBorderConverter.Convert(Common.Get(borders, cell.BorderId?.Value, cell.ApplyBorder?.Value), context, configuration));
                    styles.Styles.Merge(configuration.ConverterComposition.XlsxAlignmentConverter.Convert(Common.Get(cell.Alignment, cell.ApplyAlignment?.Value), context, configuration));
                }
                if (configuration.ConvertNumberFormats)
                {
                    styles.NumberFormatIdentifier = Common.Get(cell.NumberFormatId?.Value, cell.ApplyNumberFormat?.Value);
                }

                return styles;
            }) ?? []];
            result.DifferentialStyles = [.. value.DifferentialFormats?.Elements().Select((x, i) =>
            {
                if (x is not DifferentialFormat differential)
                {
                    return new();
                }

                Specification.Xlsx.XlsxDifferentialStyles styles = new()
                {
                    Name = configuration.UseHtmlClasses ? string.Concat("differential-", Common.Format(i, configuration)) : null,
                    IsHidden = Common.Get(differential.Protection?.Hidden?.Value, configuration.ConvertVisibilities)
                };
                if (configuration.ConvertStyles)
                {
                    styles.FontStyles = configuration.ConverterComposition.XlsxFontConverter.Convert(differential.Font, context, configuration);
                    styles.FillStyles = configuration.ConverterComposition.XlsxFillConverter.Convert(differential.Fill, context, configuration);
                    styles.BorderStyles = configuration.ConverterComposition.XlsxBorderConverter.Convert(differential.Border, context, configuration);
                    styles.AlignmentStyles = configuration.ConverterComposition.XlsxAlignmentConverter.Convert(differential.Alignment, context, configuration);
                }
                if (configuration.ConvertNumberFormats)
                {
                    styles.NumberFormat = parser(differential.NumberingFormat?.FormatCode?.Value, null);
                }

                return styles;
            }) ?? []];

            foreach (NumberingFormat number in Common.Get(value.NumberingFormats, configuration.ConvertNumberFormats)?.Elements<NumberingFormat>() ?? [])
            {
                if (number.NumberFormatId?.Value == null)
                {
                    continue;
                }

                parser(number.FormatCode?.Value, number.NumberFormatId?.Value);
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxSharedStringTableReader"/> class.
    /// </summary>
    public class DefaultXlsxSharedStringTableReader() : IConverterBase<SharedStringTable?, Specification.Xlsx.XlsxString[]>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxString[] Convert(SharedStringTable? value, ConverterContext context, ConverterConfiguration configuration)
        {
            return value != null ? [.. value.Elements().Select(x => configuration.ConverterComposition.XlsxStringConverter.Convert(x, context, configuration))] : [];
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxWorksheetReader"/> class.
    /// </summary>
    public class DefaultXlsxWorksheetReader() : IConverterBase<Worksheet?, Specification.Xlsx.XlsxSheet>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxSheet Convert(Worksheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxSheet result = new()
            {
                CellSize = (8.11, 20)
            };
            bool isDimensioned = false;

            Dictionary<uint, (double? Width, bool? IsHidden, uint? StylesIndex)> definitions = [];
            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case SheetData data:
                        result.Data = data;
                        break;
                    case SheetDimension dimension when dimension.Reference?.Value != null:
                        result.Dimension = new(dimension.Reference.Value);
                        isDimensioned = true;
                        break;
                    case SheetProperties properties when configuration.ConvertSheetTitles:
                        Specification.Html.HtmlStyles variables = [];
                        variables.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.SheetTitle, Specification.Html.HtmlStyleType.VariableSheetColorExact, configuration.ConverterComposition.XlsxColorConverter.Convert(properties.TabColor, context, configuration)), context, configuration));
                        result.TitleAttributes[Common.ATTRIBUTE_STYLE] = variables;
                        break;
                    case SheetFormatProperties format:
                        if ((format.ZeroHeight?.Value ?? false) && configuration.ConvertVisibilities)
                        {
                            result.RowAttributes[Common.ATTRIBUTE_STYLE] = new Specification.Html.HtmlStyles(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.VisibilityCollapsed), context, configuration));
                        }

                        if (configuration.ConvertStyles)
                        {
                            Specification.Html.HtmlStyles styles = [];
                            if (format.ThickTop?.Value ?? false)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.BorderThicknessTopThick), context, configuration));
                            }
                            if (format.ThickBottom?.Value ?? false)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Row, Specification.Html.HtmlStyleType.BorderThicknessBottomThick), context, configuration));
                            }

                            if (styles.Any())
                            {
                                result.CellAttributes[Common.ATTRIBUTE_STYLE] = styles;
                            }
                        }

                        if (configuration.ConvertSizes)
                        {
                            result.CellSize = (format.DefaultColumnWidth?.Value ?? format.BaseColumnWidth?.Value ?? result.CellSize.Width, (Common.Get(format.DefaultRowHeight?.Value, format.CustomHeight?.Value) * Common.RATIO_POINT) ?? result.CellSize.Height);
                        }

                        break;
                    case Columns columns:
                        foreach (Column column in columns.Elements<Column>())
                        {
                            if (column.Min?.Value == null)
                            {
                                continue;
                            }

                            for (uint i = column.Min.Value; i <= (column.Max?.Value ?? column.Min.Value); i++)
                            {
                                double? size = null;
                                if (configuration.ConvertSizes)
                                {
                                    if (column.Width?.Value != null && (column.CustomWidth?.Value ?? true))
                                    {
                                        size = column.Width.Value;
                                    }
                                    else if (column.BestFit?.Value ?? false)
                                    {
                                        size = double.NaN;
                                    }
                                }

                                definitions[i] = (size, Common.Get(column.Hidden?.Value, configuration.ConvertVisibilities), column.Style?.Value);
                            }
                        }
                        break;
                }
            }

            if (!isDimensioned)
            {
                uint column = 1;
                uint row = 1;
                foreach (Specification.Xlsx.XlsxCell cell in configuration.ConverterComposition.XlsxWorksheetIterator.Convert(result, context, configuration))
                {
                    column = Math.Max(column, cell.Reference.Column);
                    row = Math.Max(row, cell.Reference.Row);
                }

                result.Dimension.ColumnEnd = column;
                result.Dimension.RowEnd = row;
            }
            if (configuration.XlsxSheetDimensionSelector != null)
            {
                (uint left, uint top, uint right, uint bottom) = configuration.XlsxSheetDimensionSelector((result.Dimension.ColumnStart, result.Dimension.RowStart, result.Dimension.ColumnEnd, result.Dimension.RowEnd));
                result.Dimension = new(Math.Max(1, Math.Min(left, right)), Math.Max(1, Math.Min(top, bottom)), Math.Max(1, Math.Max(left, right)), Math.Max(1, Math.Max(top, bottom)));
            }

            result.Columns = new (double? Width, bool? IsHidden, uint? StylesIndex)[result.Dimension.ColumnCount];
            for (uint i = 0; i < result.Columns.Length; i++)
            {
                result.Columns[i] = Common.Get(definitions, result.Dimension.ColumnStart + i);
            }

            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case MergeCells merges:
                        foreach (MergeCell merge in merges.Elements<MergeCell>())
                        {
                            if (merge.Reference?.Value == null)
                            {
                                continue;
                            }

                            result.Specialties.Add(new(merge)
                            {
                                Range = new(merge.Reference.Value, result.Dimension)
                            });
                        }
                        break;
                    case ConditionalFormatting conditional when conditional.SequenceOfReferences?.Items != null:
                        foreach (string? item in conditional.SequenceOfReferences.Items)
                        {
                            if (item == null)
                            {
                                continue;
                            }

                            result.Specialties.Add(new(conditional)
                            {
                                Range = new(item, result.Dimension)
                            });
                        }
                        break;
                }
            }

            //TODO: support for hyperlinks

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxCellContentReader"/> class.
    /// </summary>
    public class DefaultXlsxCellContentReader() : IConverterBase<Specification.Xlsx.XlsxCell?, Specification.Xlsx.XlsxCell>
    {
        internal enum CommonStyleType
        {
            AlignmentCenter,
            AlignmentRight,
            AlignmentDistributed,
            ColorBlack,
            ColorGreen,
            ColorWhite,
            ColorBlue,
            ColorMagenta,
            ColorYellow,
            ColorCyan,
            ColorRed
        }

        internal class NumberInformation
        {
            public List<string> Tokens { get; set; } = [];
            public int Scaling { get; set; } = 0;
            public bool IsGrouped { get; set; } = false;
            public bool IsFractional { get; set; } = false;
            public int[] Lengths { get; set; } = [0, 0, 0, 0];
        }

        /// <summary>
        /// Represents the default number formats.
        /// </summary>
        protected internal static Dictionary<uint, Specification.Xlsx.XlsxNumberFormat> FORMATS = new()
        {
            [1] = new(new("0", false)),
            [2] = new(new("0.00", false)),
            [3] = new(new("#,##0", false)),
            [4] = new(new("#,##0.00", false)),
            [9] = new(new("0%", false)),
            [10] = new(new("0.00%", false)),
            [11] = new(new("0.00E+00", false)),
            [12] = new(new("# ?/?", false)),
            [13] = new(new("# ??/??", false)),
            [14] = new(new("mm-dd-yy", true)),
            [15] = new(new("d-mmm-yy", true)),
            [16] = new(new("d-mmm", true)),
            [17] = new(new("mmm-yy", true)),
            [18] = new(new("h:mm AM/PM", true)),
            [19] = new(new("h:mm:ss AM/PM", true)),
            [20] = new(new("h:mm", true)),
            [21] = new(new("h:mm:ss", true)),
            [22] = new(new("m/d/yy h:mm", true)),
            [37] = new(new("#,##0 ", false), new("(#,##0)", false)),
            [38] = new(new("#,##0 ", false), new("[Red](#,##0)", false)),
            [39] = new(new("#,##0.00", false), new("(#,##0.00)", false)),
            [40] = new(new("#,##0.00", false), new("[Red](#,##0.00)", false)),
            [45] = new(new("mm:ss", true)),
            [46] = new(new("[h]:mm:ss", true)),
            [47] = new(new("mmss.0", true)),
            [48] = new(new("##0.0E+0", false)),
            [49] = new(new("@", false))
        };

        /// <summary>
        /// Represents the colors of a number format.
        /// </summary>
        protected internal static Dictionary<string, object> COLORS = new()
        {
            ["BLACK"] = CommonStyleType.ColorBlack,
            ["GREEN"] = CommonStyleType.ColorGreen,
            ["WHITE"] = CommonStyleType.ColorWhite,
            ["BLUE"] = CommonStyleType.ColorBlue,
            ["MAGENTA"] = CommonStyleType.ColorMagenta,
            ["YELLOW"] = CommonStyleType.ColorYellow,
            ["CYAN"] = CommonStyleType.ColorCyan,
            ["RED"] = CommonStyleType.ColorRed
        };

        /// <summary>
        /// Represents the conditions of a number format.
        /// </summary>
        protected internal static Dictionary<string, Func<double, double, bool>> CONDITIONS = new()
        {
            ["="] = (x, y) => x == y,
            ["<>"] = (x, y) => x != y,
            ["<"] = (x, y) => x < y,
            ["<="] = (x, y) => x <= y,
            [">"] = (x, y) => x > y,
            [">="] = (x, y) => x >= y
        };

        /// <inheritdoc />
        public Specification.Xlsx.XlsxCell Convert(Specification.Xlsx.XlsxCell? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new(null);
            }

            Specification.Xlsx.XlsxBaseStyles stylizer(CommonStyleType type)
            {
                object key = (Common.CacheCategory.CommonStyles, type);
                if (Common.Get(context.Cache, key) is not Specification.Xlsx.XlsxBaseStyles cache)
                {
                    Specification.Html.HtmlStyles styles = [];
                    switch (type)
                    {
                        case CommonStyleType.AlignmentCenter:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.AlignmentHorizontalCenter), context, configuration));
                            break;
                        case CommonStyleType.AlignmentRight:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.AlignmentHorizontalRight), context, configuration));
                            break;
                        case CommonStyleType.AlignmentDistributed:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.AlignmentHorizontalDistributed), context, configuration));
                            break;
                        case CommonStyleType.ColorBlack:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "black"), context, configuration));
                            break;
                        case CommonStyleType.ColorGreen:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "green"), context, configuration));
                            break;
                        case CommonStyleType.ColorWhite:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "white"), context, configuration));
                            break;
                        case CommonStyleType.ColorBlue:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "blue"), context, configuration));
                            break;
                        case CommonStyleType.ColorMagenta:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "magenta"), context, configuration));
                            break;
                        case CommonStyleType.ColorYellow:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "yellow"), context, configuration));
                            break;
                        case CommonStyleType.ColorCyan:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "cyan"), context, configuration));
                            break;
                        case CommonStyleType.ColorRed:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.ForegroundExact, "red"), context, configuration));
                            break;
                    }

                    cache = new()
                    {
                        Styles = new(styles)
                    };
                    context.Cache[key] = cache;
                }

                return cache;
            }
            (string Raw, Specification.Html.HtmlChildren Children) disassembler(Specification.Xlsx.XlsxString data)
            {
                return (data.Raw, data.Children);
            }
            Specification.Html.HtmlChildren formatter(object data, string raw)
            {
                if (!configuration.ConvertNumberFormats)
                {
                    return [raw];
                }

                Specification.Html.HtmlChildren wrapper(object data, Specification.Html.HtmlChildren children)
                {
                    return configuration.UseHtmlDataElements ? data switch
                    {
                        DateTime date => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "time", new()
                        {
                            ["datetime"] = date.ToString("yyyy-MM-ddThh:mm:ss.fff", CultureInfo.InvariantCulture)
                        }, children)],
                        double number => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "data", new()
                        {
                            ["value"] = number.ToString(CultureInfo.InvariantCulture)
                        }, children)],
                        _ => children
                    } : children;
                }
                List<string> tokenizer(string code, bool isStandardized, Func<char, StringBuilder, bool?> parser, char[]? singles)
                {
                    bool isSpecial = false;
                    StringBuilder builder = new();
                    List<string> tokens = [];
                    foreach ((int index, char character, bool isEscaped) in Specification.Xlsx.XlsxNumberFormat.Escape(code, singles))
                    {
                        if (isEscaped)
                        {
                            builder.Append(character);
                            continue;
                        }

                        char input = isStandardized ? char.ToUpperInvariant(character) : character;
                        bool? isInitial = parser(input, builder);
                        if (isInitial ?? isSpecial)
                        {
                            tokens.Add(builder.ToString());
                            builder.Clear();
                        }

                        isSpecial = isInitial != null;
                        builder.Append(isSpecial ? input : character);
                    }
                    tokens.Add(builder.ToString());

                    return tokens;
                }
                void escaper(StringBuilder builder, string content)
                {
                    foreach ((int index, char character, bool isEscaped) in Specification.Xlsx.XlsxNumberFormat.Escape(content))
                    {
                        if (!isEscaped && character is '\\' or '\"')
                        {
                            continue;
                        }

                        builder.Append(character);
                    }
                }

                //TODO: support for locale-dependent default formats

                (int section, Specification.Xlsx.XlsxNumberFormatCode? code) = (value.NumberFormat ?? (value.NumberFormatIdentifier != null ? Common.Get(FORMATS, value.NumberFormatIdentifier.Value) : null)) is Specification.Xlsx.XlsxNumberFormat format ? data switch
                {
                    double positive when positive > 0 => (0, format.Positive),
                    double negative when negative < 0 => (1, format.Negative),
                    double zero when zero == 0 => (2, format.Zero),
                    _ => (3, format.Text)
                } : (-1, null);
                object? key = value.NumberFormatIdentifier != null ? (Common.CacheCategory.NumberFormat, value.NumberFormatIdentifier.Value, section) : null;

                string? currency = null;
                CultureInfo culture = configuration.CurrentCulture;
                if (code != null)
                {
                    int start = 0;
                    Specification.Xlsx.XlsxStyles? styles = null;

                    string? token = null;
                    foreach ((int index, char character, bool isEscaped) in Specification.Xlsx.XlsxNumberFormat.Escape(code.Code))
                    {
                        if (token == null)
                        {
                            if (!isEscaped && character is '[')
                            {
                                token = string.Empty;
                            }
                            else if (!char.IsWhiteSpace(character))
                            {
                                break;
                            }

                            continue;
                        }
                        else if (!isEscaped && character is not ']')
                        {
                            token += char.ToUpperInvariant(character);
                            continue;
                        }

                        if (token.StartsWith('$'))
                        {
                            string[] identifiers = token.TrimStart('$').Split('-');
                            currency = !identifiers[0].All(char.IsWhiteSpace) ? identifiers[0] : null;
                            if (Common.ParseHex(identifiers[^1]) is int locale)
                            {
                                try
                                {
                                    culture = locale switch
                                    {
                                        0xF800 or 0xF400 => configuration.CurrentCulture,
                                        _ => CultureInfo.GetCultureInfo(locale)
                                    };
                                }
                                catch { }
                            }
                        }
                        else if (Common.Get(COLORS, token) is CommonStyleType color)
                        {
                            if (configuration.ConvertStyles)
                            {
                                styles = stylizer(color);
                            }
                        }
                        else if (Common.Get(CONDITIONS, new string([.. token.TakeWhile(x => x is '=' or '<' or '>')])) is Func<double, double, bool> comparator)
                        {
                            if (data is double number && Common.ParseDecimals(new string([.. token.SkipWhile(x => x is '=' or '<' or '>')])) is double operand && comparator(number, operand))
                            {
                                styles = null;
                            }
                        }
                        else
                        {
                            break;
                        }

                        token = null;
                        start = index + 1;
                    }

                    if (styles != null)
                    {
                        value.Styles.Add(styles);
                    }

                    code = new(code.Code[start..], code.IsCalendrical);
                }

                if (code == null || code.Code.All(char.IsWhiteSpace) || code.Code.Trim().ToUpperInvariant() == "GENERAL")
                {
                    switch (data)
                    {
                        case DateTime date:
                            return wrapper(date, [date.ToString("d", culture)]);
                        case double number:
                            string general = number.ToString(CultureInfo.InvariantCulture).Replace("+", string.Empty);
                            if (general.Length <= (general.StartsWith('-') ? 12 : 11))
                            {
                                return wrapper(number, [general]);
                            }

                            string scientific = number.ToString("0.#######E0", CultureInfo.InvariantCulture);

                            return wrapper(number, [number.ToString(string.Concat("0.", new string('#', Math.Max(0, (scientific.StartsWith('-') ? 10 : 9) - (scientific.Length - scientific.IndexOf('E')))), "E0"), CultureInfo.InvariantCulture)]);
                        default:
                            return [raw];
                    }
                }

                StringBuilder builder = new();
                Specification.Html.HtmlChildren children = [];
                if (code.IsCalendrical)
                {
                    if (data is double number && number >= -657435.0 && number <= 2958465.99999999)
                    {
                        data = DateTime.FromOADate(number);
                    }
                    if (data is not DateTime date)
                    {
                        return [raw];
                    }

                    if (Common.Get(context.Cache, key) is not List<string> information)
                    {
                        information = tokenizer(code.Code, true, (x, y) => x switch
                        {
                            '[' or 'A' => y.Length > 0,
                            ']' when y.Length > 0 && y[0] is '[' => false,
                            'M' when (y.Length == 1 && y[0] is 'A') || (y.Length == 4 && y[0] is 'A' && y[1] is 'M' && y[2] is '/' && y[3] is 'P') => false,
                            '/' when (y.Length == 1 && y[0] is 'A') || (y.Length == 2 && y[0] is 'A' && y[1] is 'M') => false,
                            'P' when (y.Length == 2 && y[0] is 'A' && y[1] is '/') || (y.Length == 3 && y[0] is 'A' && y[1] is 'M' && y[2] is '/') => false,
                            '.' when y.Length > 0 && y[0] is 'S' && y[^1] is 'S' => false,
                            '0' or '#' when y.Length > 0 && y[^1] is '.' or '0' or '#' => false,
                            'Y' or 'M' or 'D' or 'H' or 'S' => y.Length > 0 && y[0] != x && y[0] is not '[',
                            '@' or '$' or '/' or ':' => true,
                            _ => null
                        }, null);

                        if (key != null)
                        {
                            context.Cache[key] = information;
                        }
                    }

                    bool determiner(int index)
                    {
                        (int Distance, bool? IsTemporal) left = (0, null);
                        (int Distance, bool? IsTemporal) right = (0, null);
                        for (int i = 1; index - i >= 0 && left.IsTemporal == null; i++)
                        {
                            left = information[index - i].FirstOrDefault(char.IsLetter) switch
                            {
                                'H' or 'S' => (left.Distance, true),
                                'Y' or 'D' => (left.Distance, false),
                                _ => (left.Distance + information[index - i].Length, null)
                            };
                        }
                        for (int i = 1; index + i < information.Count && right.IsTemporal == null && right.Distance <= left.Distance; i++)
                        {
                            right = information[index + i].FirstOrDefault(char.IsLetter) switch
                            {
                                'H' or 'S' => (right.Distance, true),
                                'Y' or 'D' => (right.Distance, false),
                                _ => (right.Distance + information[index + i].Length, null)
                            };
                        }

                        return (left.IsTemporal != right.IsTemporal && left.Distance > right.Distance ? right.IsTemporal : left.IsTemporal) ?? false;
                    }
                    TimeSpan measurer()
                    {
                        return date.Year < 100 || date.Year > 9999 ? TimeSpan.Zero : TimeSpan.FromDays(date.ToOADate());
                    }

                    bool isDivided = information.Any(x => x == "A/P" || x == "AM/PM");
                    for (int i = 0; i < information.Count; i++)
                    {
                        string token = information[i];

                        string suffix = string.Empty;
                        if (token.Contains('.'))
                        {
                            string[] parts = token.Split('.');
                            if (parts[^1].Any() && parts[^1].All(x => x is '0' or '#'))
                            {
                                token = parts[0];
                                suffix = string.Concat(".", date.Millisecond.ToString(parts[^1], culture));
                            }
                        }

                        if (token switch
                        {
                            "@" => raw,
                            "$" => currency ?? culture.NumberFormat.CurrencySymbol,
                            "/" => culture.DateTimeFormat.DateSeparator,
                            ":" => culture.DateTimeFormat.TimeSeparator,
                            "YY" => date.ToString("yy", culture),
                            "YYYY" => date.ToString("yyyy", culture),
                            "M" => date.ToString(determiner(i) ? "%m" : "%M", culture),
                            "MM" => date.ToString(determiner(i) ? "mm" : "MM", culture),
                            "MMM" => date.ToString("MMM", culture),
                            "MMMM" => date.ToString("MMMM", culture),
                            "MMMMM" => date.ToString("MMMM", culture).FirstOrDefault().ToString(CultureInfo.InvariantCulture),
                            "D" => date.ToString("%d", culture),
                            "DD" => date.ToString("dd", culture),
                            "DDD" => date.ToString("ddd", culture),
                            "DDDD" => date.ToString("dddd", culture),
                            "H" => date.ToString(isDivided ? "%h" : "%H", culture),
                            "HH" => date.ToString(isDivided ? "hh" : "HH", culture),
                            "S" => date.ToString("%s", culture),
                            "SS" => date.ToString("ss", culture),
                            "A/P" => date.ToString("%t", culture),
                            "AM/PM" => date.ToString("tt", culture),
                            "[H]" => measurer().TotalHours.ToString("0", culture),
                            "[M]" => measurer().TotalMinutes.ToString("0", culture),
                            "[MM]" => measurer().TotalMinutes.ToString("00", culture),
                            "[S]" => measurer().TotalSeconds.ToString("0", culture),
                            "[SS]" => measurer().TotalSeconds.ToString("00", culture),
                            _ => null
                        } is string content)
                        {
                            builder.Append(content);
                        }
                        else
                        {
                            escaper(builder, token);
                        }

                        builder.Append(suffix);
                    }
                }
                else
                {
                    double number = Math.Abs(data as double? ?? 0);

                    int stage = 0;
                    if (Common.Get(context.Cache, key) is not NumberInformation information)
                    {
                        information = new();
                        information.Tokens = tokenizer(code.Code, false, (x, y) =>
                        {
                            switch (x)
                            {
                                case '0' or '#' or '?':
                                    information.Lengths[stage]++;
                                    if (y.Length > 0 && y[0] is ',')
                                    {
                                        information.Scaling += 3;
                                        information.IsGrouped = true;
                                    }

                                    return y.Length > 0 && (y[0] is '_' or '*' || !(y[^1] is '0' or '#' or '?' or '/'));
                                case '.' when stage < 1:
                                    stage = 1;
                                    return true;
                                case ',':
                                    information.Scaling -= 3;
                                    return true;
                                case 'E' or 'e' when stage < 2:
                                    stage = 2;
                                    return true;
                                case '+' or '-' when y.Length == 1 && y[0] is 'E' or 'e':
                                    return false;
                                case '%':
                                    information.Scaling += 2;
                                    return true;
                                case '/' when stage < 1 && y.Length > 0 && y[^1] is '0' or '#' or '?':
                                    stage = 3;
                                    information.IsFractional = true;
                                    information.Lengths[0] -= y.Length;
                                    return false;
                                case '@' or '$' or '_' or '*':
                                    return true;
                                default:
                                    return y.Length == 1 && y[0] is '_' or '*' ? false : null;
                            }
                        }, ['_', '*']);

                        if (key != null)
                        {
                            context.Cache[key] = information;
                        }
                    }

                    number *= Math.Pow(10, information.Scaling);

                    string numerator = string.Empty;
                    string denominator = string.Empty;
                    if (information.IsFractional)
                    {
                        long whole = (long)number;
                        double remainder = number - whole;
                        (int Numerator, int Denominator)? fraction = remainder switch
                        {
                            < 0.001 => (0, 1),
                            > 0.999 => (1, 1),
                            _ => null
                        };

                        int maximum = (int)Math.Pow(10, Math.Min(4, information.Lengths[3])) - 1;
                        (int Numerator, int Denominator) lower = (0, 1);
                        (int Numerator, int Denominator) upper = (1, 1);
                        while (fraction == null)
                        {
                            (int Numerator, int Denominator) middle = (lower.Numerator + upper.Numerator, lower.Denominator + upper.Denominator);
                            if (middle.Denominator > maximum)
                            {
                                fraction = Math.Abs(remainder - (double)lower.Numerator / lower.Denominator) <= Math.Abs(remainder - (double)upper.Numerator / upper.Denominator) ? lower : upper;
                            }
                            else if (middle.Numerator < middle.Denominator * (remainder - 0.001))
                            {
                                lower = middle;
                            }
                            else if (middle.Numerator > middle.Denominator * (remainder + 0.001))
                            {
                                upper = middle;
                            }
                            else
                            {
                                fraction = (middle.Numerator, middle.Denominator);
                            }
                        }

                        if (information.Lengths[0] <= 0)
                        {
                            number = 0;
                            numerator = (fraction.Value.Denominator * whole + fraction.Value.Numerator).ToString(CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            number = whole;
                            numerator = fraction.Value.Numerator.ToString(CultureInfo.InvariantCulture);
                        }

                        denominator = fraction.Value.Denominator.ToString(CultureInfo.InvariantCulture);
                    }

                    char sign = '+';
                    string[] components = [string.Empty, string.Empty, string.Empty];
                    if (information.Lengths[2] > 0)
                    {
                        int exponent = number > 0 ? (int)Math.Floor(Math.Log10(number)) : 0;
                        number = Math.Round(number / Math.Pow(10, exponent), information.Lengths[1]);
                        if (number >= 10)
                        {
                            number /= 10;
                            exponent++;
                        }

                        sign = exponent < 0 ? '-' : '+';
                        components[2] = exponent.ToString(CultureInfo.InvariantCulture).PadLeft(information.Lengths[2], ' ');
                    }
                    else
                    {
                        number = Math.Round(number, information.Lengths[1]);
                    }

                    long integer = (long)number;
                    components[0] = integer.ToString(CultureInfo.InvariantCulture).PadLeft(information.Lengths[0], ' ');
                    components[1] = information.Lengths[1] > 0 ? (number - integer).ToString(string.Concat(".", new string('#', information.Lengths[1])), CultureInfo.InvariantCulture).TrimStart('.').PadRight(information.Lengths[1], ' ') : string.Empty;

                    List<int> separators = [];
                    if (information.IsGrouped && culture.NumberFormat.NumberGroupSizes.Any())
                    {
                        int group = 0;
                        int length = components[0].Length - 1;
                        while (length > 0)
                        {
                            int size = culture.NumberFormat.NumberGroupSizes[Math.Min(culture.NumberFormat.NumberGroupSizes.Length - 1, group)];
                            if (size <= 0)
                            {
                                break;
                            }

                            length -= size;
                            separators.Add(length);
                            group++;
                        }
                    }

                    int concatenator(string token, string source, int start)
                    {
                        int index = start;
                        foreach (char character in token)
                        {
                            builder.Append(source[index] is ' ' ? character switch
                            {
                                '0' => '0',
                                '?' => ' ',
                                _ => string.Empty
                            } : (stage < 1 && separators.Contains(index) ? string.Concat(source[index].ToString(CultureInfo.InvariantCulture), culture.NumberFormat.NumberGroupSeparator) : source[index]));

                            index++;
                        }

                        return index;
                    }

                    stage = 0;
                    int index = 0;
                    for (int i = 0; i < information.Tokens.Count; i++)
                    {
                        string token = information.Tokens[i];
                        switch (token.FirstOrDefault())
                        {
                            case '@':
                                builder.Append(raw);
                                break;
                            case '$':
                                builder.Append(currency ?? culture.NumberFormat.CurrencySymbol);
                                break;
                            case '0' or '#' or '?' when stage < 3:
                                if (information.IsFractional && token.Contains('/'))
                                {
                                    string[] parts = token.Split('/');
                                    string left = parts[0];
                                    string right = parts[^1];
                                    concatenator(left.PadLeft(numerator.Length, '0'), numerator.PadLeft(left.Length, ' '), 0);
                                    builder.Append('/');
                                    concatenator(right.PadRight(denominator.Length, '0'), denominator.PadRight(right.Length, ' '), 0);

                                    stage = 3;

                                    break;
                                }

                                if (stage != 1 && index <= 0)
                                {
                                    index = concatenator(new('0', components[stage].Length - information.Lengths[stage]), components[stage], index);
                                }

                                index = concatenator(token, components[stage], index);

                                break;
                            case '.' when stage < 1:
                                if (index <= 0)
                                {
                                    index = concatenator(new('0', components[0].Length - information.Lengths[0]), components[0], index);
                                }

                                stage = 1;
                                index = 0;
                                builder.Append(culture.NumberFormat.NumberDecimalSeparator);

                                break;
                            case ',':
                                break;
                            case 'E' or 'e' when stage < 2:
                                stage = 2;
                                index = 0;
                                builder.Append(sign is '-' || token.Length > 1 ? new string([token[0], sign]) : token[0]);
                                break;
                            case '%':
                                builder.Append(culture.NumberFormat.PercentSymbol);
                                break;
                            case '_':
                                builder.Append(' ');
                                break;
                            case '*':
                                children.Add(builder.ToString());
                                builder.Clear();
                                break;
                            default:
                                escaper(builder, token);
                                break;
                        }
                    }
                }
                children.Add(builder.ToString());

                if (children.Count > 1)
                {
                    Specification.Html.HtmlElement container = new(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, null, [.. children.Select(x => new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, null, [x]))]);
                    stylizer(CommonStyleType.AlignmentDistributed).ApplyStyles(container);
                    children = [container];
                }

                return wrapper(data, children);
            }

            string content = value.Cell?.CellValue?.Text ?? string.Empty;
            (string raw, Specification.Html.HtmlChildren? children) = value.Cell?.DataType?.Value switch
            {
                _ when value.Cell?.DataType?.Value == CellValues.Error => (content, [content]),
                _ when value.Cell?.DataType?.Value == CellValues.String => (content, [content]),
                _ when value.Cell?.DataType?.Value == CellValues.InlineString => disassembler(configuration.ConverterComposition.XlsxStringConverter.Convert(value.Cell, context, configuration)),
                _ when value.Cell?.DataType?.Value == CellValues.SharedString => Common.ParsePositive(content) is uint index && Common.Get(context.SharedStrings, index) is Specification.Xlsx.XlsxString shared ? disassembler(shared) : (string.Empty, []),
                _ when value.Cell?.DataType?.Value == CellValues.Boolean => (content, [content.Trim() switch
                {
                    "1" => "TRUE",
                    "0" => "FALSE",
                    _ => string.Empty
                }]),
                _ => (content, null)
            };

            foreach (Specification.Xlsx.XlsxSpecialty specialty in value.Specialties)
            {
                switch (specialty.Specialty)
                {
                    case ConditionalFormatting conditional:
                        foreach (ConditionalFormattingRule rule in conditional.Elements<ConditionalFormattingRule>().OrderByDescending(x => x.Priority?.Value ?? int.MaxValue))
                        {
                            if (rule.Type?.Value == null)
                            {
                                continue;
                            }

                            bool comparator(ConditionalFormattingOperatorValues operation)
                            {
                                double? number = Common.ParseDecimals(raw);
                                string[] formulas = [.. rule.Elements<Formula>().Select(x => WebUtility.HtmlDecode(x.Text.Trim('\"')))];
                                double?[] operands = [.. formulas.Select(Common.ParseDecimals)];

                                return operation switch
                                {
                                    _ when operation == ConditionalFormattingOperatorValues.Equal && formulas.Length > 0 => raw.Equals(formulas[0], StringComparison.OrdinalIgnoreCase) || (number != null && operands[0] != null && number == operands[0]),
                                    _ when operation == ConditionalFormattingOperatorValues.NotEqual && formulas.Length > 0 => !raw.Equals(formulas[0], StringComparison.OrdinalIgnoreCase) || (number != null && operands[0] != null && number != operands[0]),
                                    _ when operation == ConditionalFormattingOperatorValues.LessThan && formulas.Length > 0 && number != null && operands[0] != null => number < operands[0],
                                    _ when operation == ConditionalFormattingOperatorValues.LessThanOrEqual && formulas.Length > 0 && number != null && operands[0] != null => number <= operands[0],
                                    _ when operation == ConditionalFormattingOperatorValues.GreaterThan && formulas.Length > 0 && number != null && operands[0] != null => number > operands[0],
                                    _ when operation == ConditionalFormattingOperatorValues.GreaterThanOrEqual && formulas.Length > 0 && number != null && operands[0] != null => number >= operands[0],
                                    _ when operation == ConditionalFormattingOperatorValues.Between && formulas.Length > 1 && number != null && operands[0] != null && operands[1] != null => number >= Math.Min(operands[0] ?? 0, operands[1] ?? 0) && number <= Math.Max(operands[0] ?? 0, operands[1] ?? 0),
                                    _ when operation == ConditionalFormattingOperatorValues.NotBetween && formulas.Length > 1 && number != null && operands[0] != null && operands[1] != null => number < Math.Min(operands[0] ?? 0, operands[1] ?? 0) || number > Math.Max(operands[0] ?? 0, operands[1] ?? 0),
                                    _ when operation == ConditionalFormattingOperatorValues.ContainsText && formulas.Length > 0 => raw.Contains(formulas[0], StringComparison.OrdinalIgnoreCase),
                                    _ when operation == ConditionalFormattingOperatorValues.NotContains && formulas.Length > 0 => !raw.Contains(formulas[0], StringComparison.OrdinalIgnoreCase),
                                    _ when operation == ConditionalFormattingOperatorValues.BeginsWith && formulas.Length > 0 => raw.StartsWith(formulas[0], StringComparison.OrdinalIgnoreCase),
                                    _ when operation == ConditionalFormattingOperatorValues.EndsWith && formulas.Length > 0 => raw.EndsWith(formulas[0], StringComparison.OrdinalIgnoreCase),
                                    _ => false
                                };
                            }

                            //TODO: support for more conditional formatting rules

                            if (rule.Type.Value switch
                            {
                                _ when rule.Type.Value == ConditionalFormatValues.CellIs && rule.Operator?.Value != null => comparator(rule.Operator.Value),
                                _ when rule.Type.Value == ConditionalFormatValues.ContainsText && rule.Text?.Value != null => raw.Contains(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.NotContainsText && rule.Text?.Value != null => !raw.Contains(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.BeginsWith && rule.Text?.Value != null => raw.StartsWith(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.EndsWith && rule.Text?.Value != null => raw.EndsWith(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.ContainsBlanks => raw.Any(char.IsWhiteSpace),
                                _ when rule.Type.Value == ConditionalFormatValues.NotContainsBlanks => !raw.Any(char.IsWhiteSpace),
                                _ when rule.Type.Value == ConditionalFormatValues.ContainsErrors => value.Cell?.DataType?.Value == CellValues.Error,
                                _ when rule.Type.Value == ConditionalFormatValues.NotContainsErrors => value.Cell?.DataType?.Value != CellValues.Error,
                                _ => false,
                            })
                            {
                                if (Common.Get(context.Stylesheet.DifferentialStyles, rule.FormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles differential)
                                {
                                    value.Styles.Add(differential);
                                    if (differential.NumberFormat != null && configuration.ConvertNumberFormats)
                                    {
                                        value.NumberFormat = differential.NumberFormat;
                                    }
                                }

                                if (rule.StopIfTrue?.Value ?? false)
                                {
                                    break;
                                }
                            }
                        }

                        break;
                }
            }

            if (children == null)
            {
                (children, bool isAligned) = value.Cell?.DataType?.Value switch
                {
                    _ when value.Cell?.DataType?.Value == CellValues.Date => DateTime.TryParse(content, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date) ? (formatter(date, content), true) : ([content], false),
                    _ => Common.ParseDecimals(content) is double number ? (formatter(number, content), true) : (formatter(content, content), false)
                };

                if (isAligned && configuration.ConvertStyles)
                {
                    value.Styles.Insert(0, stylizer(CommonStyleType.AlignmentRight));
                }
            }
            else if ((value.Cell?.DataType?.Value == CellValues.Error || value.Cell?.DataType?.Value == CellValues.Boolean) && configuration.ConvertStyles)
            {
                value.Styles.Insert(0, stylizer(CommonStyleType.AlignmentCenter));
            }

            value.Children = children;

            return value;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxTableReader"/> class.
    /// </summary>
    public class DefaultXlsxTableReader() : IConverterBase<TableDefinitionPart?, IEnumerable<Specification.Xlsx.XlsxSpecialty>>
    {
        /// <inheritdoc />
        public IEnumerable<Specification.Xlsx.XlsxSpecialty> Convert(TableDefinitionPart? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value?.Table == null || value.Table.Reference?.Value == null)
            {
                yield break;
            }

            Specification.Xlsx.XlsxDifferentialStyles? stylizer(uint? content, uint? border)
            {
                Specification.Xlsx.XlsxDifferentialStyles? styles = null;
                if (Common.Get(context.Stylesheet.DifferentialStyles, content) is Specification.Xlsx.XlsxDifferentialStyles body)
                {
                    if (body.IsHidden != null && configuration.ConvertVisibilities)
                    {
                        styles ??= new();
                        styles.IsHidden = body.IsHidden;
                    }
                    if (configuration.ConvertStyles)
                    {
                        styles ??= new();
                        styles.FontStyles = body.FontStyles;
                        styles.FillStyles = body.FillStyles;
                        styles.AlignmentStyles = body.AlignmentStyles;
                    }
                    if (configuration.ConvertNumberFormats)
                    {
                        styles ??= new();
                        styles.NumberFormat = body.NumberFormat;
                    }
                }
                if (Common.Get(context.Stylesheet.DifferentialStyles, border) is Specification.Xlsx.XlsxDifferentialStyles boundary && configuration.ConvertStyles)
                {
                    styles ??= new();
                    styles.BorderStyles = boundary.BorderStyles;
                }

                return styles;
            }

            //TODO: support for table styles

            Specification.Xlsx.XlsxRange range = new(value.Table.Reference.Value, context.Sheet.Dimension);
            uint start = value.Table.HeaderRowCount?.Value ?? 1;
            uint end = value.Table.TotalsRowCount?.Value ?? 0;
            uint middle = range.RowCount - start - end;
            if (start > 0 && stylizer(value.Table.HeaderRowFormatId?.Value, value.Table.HeaderRowBorderFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles header)
            {
                yield return new(header)
                {
                    Range = new(range.ColumnStart, range.RowStart, range.ColumnEnd, range.RowStart + start - 1)
                };
            }
            if (middle > 0 && stylizer(value.Table.DataFormatId?.Value, value.Table.BorderFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles data)
            {
                yield return new(data)
                {
                    Range = new(range.ColumnStart, range.RowStart + start, range.ColumnEnd, range.RowEnd - end)
                };
            }
            if (end > 0 && stylizer(value.Table.TotalsRowFormatId?.Value, value.Table.TotalsRowBorderFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles totals)
            {
                yield return new(totals)
                {
                    Range = new(range.ColumnStart, range.RowEnd - end + 1, range.ColumnEnd, range.RowEnd)
                };
            }

            uint shift = 0;
            foreach (TableColumn column in value.Table.TableColumns?.Elements<TableColumn>() ?? [])
            {
                uint index = range.ColumnStart + shift;
                if (index > range.ColumnEnd)
                {
                    break;
                }

                if (start > 0 && Common.Get(context.Stylesheet.DifferentialStyles, column.HeaderRowDifferentialFormattingId?.Value) is Specification.Xlsx.XlsxDifferentialStyles top)
                {
                    yield return new(top)
                    {
                        Range = new(index, range.RowStart, index, range.RowStart + start - 1)
                    };
                }
                if (middle > 0 && Common.Get(context.Stylesheet.DifferentialStyles, column.DataFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles center)
                {
                    yield return new(center)
                    {
                        Range = new(index, range.RowStart + start, index, range.RowEnd - end)
                    };
                }
                if (end > 0 && Common.Get(context.Stylesheet.DifferentialStyles, column.TotalsRowDifferentialFormattingId?.Value) is Specification.Xlsx.XlsxDifferentialStyles bottom)
                {
                    yield return new(bottom)
                    {
                        Range = new(index, range.RowEnd - end + 1, index, range.RowEnd)
                    };
                }

                shift++;
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxDrawingReader"/> class.
    /// </summary>
    public class DefaultXlsxDrawingReader : IConverterBase<DrawingsPart?, IEnumerable<Specification.Xlsx.XlsxSpecialty>>
    {
        /// <inheritdoc />
        public IEnumerable<Specification.Xlsx.XlsxSpecialty> Convert(DrawingsPart? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                yield break;
            }

            Specification.Html.HtmlElement stylizer(Specification.Html.HtmlElement element, Specification.Html.HtmlStyles styles, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeStyle? shape, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties? properties, bool? isHidden)
            {
                element.Attributes[Common.ATTRIBUTE_STYLE] = styles;

                if (shape != null)
                {
                    if (shape.FontReference != null)
                    {
                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.ForegroundExact, configuration.ConverterComposition.XlsxColorConverter.Convert(shape.FontReference, context, configuration)), context, configuration));
                    }
                    if (shape.FillReference != null)
                    {
                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BackgroundExact, configuration.ConverterComposition.XlsxColorConverter.Convert(shape.FillReference, context, configuration)), context, configuration));
                    }
                    if (shape.LineReference != null)
                    {
                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BorderAllRegular, configuration.ConverterComposition.XlsxColorConverter.Convert(shape.LineReference, context, configuration)), context, configuration));
                    }
                }

                if (properties?.BlackWhiteMode?.Value != null)
                {
                    if (properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Hidden)
                    {
                        isHidden = true;
                    }
                    else if (properties.BlackWhiteMode.Value switch
                    {
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Black => "grayscale(1) brightness(0)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.BlackGray => "grayscale(1) contrast(1.4)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.BlackWhite => "grayscale(1) contrast(10) brightness(1.2)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Gray => "grayscale(1)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.LightGray => "grayscale(1) contrast(0.7) brightness(1.4)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.InvGray => "grayscale(1) invert(1)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.GrayWhite => "grayscale(1) contrast(0.5) brightness(1.7)",
                        _ when properties.BlackWhiteMode.Value == DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.White => "grayscale(1) brightness(10)",
                        _ => null
                    } is string filter)
                    {
                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.FilterExact, filter), context, configuration));
                    }
                }

                foreach (OpenXmlElement child in properties?.Elements() ?? [])
                {
                    switch (child)
                    {
                        case DocumentFormat.OpenXml.Drawing.Transform2D transform:
                            if (transform.Offset?.X?.Value != null)
                            {
                                double offset = transform.Offset.X.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationLeftExact, string.Concat(Common.Format(offset, configuration), "px")), context, configuration), true);
                                if (transform.Extents?.Cx?.Value != null)
                                {
                                    styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationRightExact, Common.Format("calc(100% - {0}px)", [Common.Format(offset + transform.Extents.Cx.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)])), context, configuration), true);
                                }
                            }
                            if (transform.Offset?.Y?.Value != null)
                            {
                                double offset = transform.Offset.Y.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationTopExact, string.Concat(Common.Format(offset, configuration), "px")), context, configuration), true);
                                if (transform.Extents?.Cy?.Value != null)
                                {
                                    styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationBottomExact, Common.Format("calc(100% - {0}px)", [Common.Format(offset + transform.Extents.Cy.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)])), context, configuration), true);
                                }
                            }
                            if (transform.Rotation?.Value != null)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.RotationNumeric, Common.Format(transform.Rotation.Value * Common.RATIO_ANGLE, configuration)), context, configuration));
                            }
                            if ((transform.HorizontalFlip?.Value ?? false) || (transform.VerticalFlip?.Value ?? false))
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.ScalingExact, string.Concat((transform.HorizontalFlip?.Value ?? false) ? "-1" : "1", " ", (transform.VerticalFlip?.Value ?? false) ? "-1" : "1")), context, configuration));
                            }

                            break;
                        case DocumentFormat.OpenXml.Drawing.NoFill:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BackgroundNone), context, configuration));
                            break;
                        case DocumentFormat.OpenXml.Drawing.SolidFill background:
                            styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BackgroundExact, configuration.ConverterComposition.XlsxColorConverter.Convert(background, context, configuration)), context, configuration));
                            break;
                        case DocumentFormat.OpenXml.Drawing.Outline outline:
                            if (outline.Width?.Value != null)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BorderThicknessAllNumeric, Common.Format(outline.Width.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)), context, configuration));
                            }
                            if (outline.CompoundLineType?.Value != null && outline.CompoundLineType.Value != DocumentFormat.OpenXml.Drawing.CompoundLineValues.Single)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BorderStyleAllDouble), context, configuration));
                            }

                            foreach (OpenXmlElement component in outline)
                            {
                                switch (component)
                                {
                                    case DocumentFormat.OpenXml.Drawing.PresetDash preset:
                                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, preset.Val?.Value switch
                                        {
                                            _ when preset.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid => Specification.Html.HtmlStyleType.BorderStyleAllSolid,
                                            _ when preset.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot => Specification.Html.HtmlStyleType.BorderStyleAllDotted,
                                            _ when preset.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.SystemDashDotDot => Specification.Html.HtmlStyleType.BorderStyleAllDotted,
                                            _ => Specification.Html.HtmlStyleType.BorderStyleAllDashed,
                                        }), context, configuration));
                                        break;
                                    case DocumentFormat.OpenXml.Drawing.CustomDash:
                                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BorderStyleAllDashed), context, configuration));
                                        break;
                                    case DocumentFormat.OpenXml.Drawing.NoFill:
                                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BackgroundNone), context, configuration));
                                        break;
                                    case DocumentFormat.OpenXml.Drawing.SolidFill border:
                                        styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BorderColorAllExact, configuration.ConverterComposition.XlsxColorConverter.Convert(border, context, configuration)), context, configuration));
                                        break;
                                }
                            }

                            break;
                        case DocumentFormat.OpenXml.Drawing.PresetGeometry preset:
                            //TODO: support for preset shapes
                            break;
                        case DocumentFormat.OpenXml.Drawing.CustomGeometry custom:
                            if (custom.Rectangle?.Top?.Value != null)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.MarginTopExact, string.Concat(Common.Format((Common.ParseLarge(custom.Rectangle.Top.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration), "px")), context, configuration));
                            }
                            if (custom.Rectangle?.Right?.Value != null)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.MarginRightExact, string.Concat(Common.Format((Common.ParseLarge(custom.Rectangle.Right.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration), "px")), context, configuration));
                            }
                            if (custom.Rectangle?.Bottom?.Value != null)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.MarginBottomExact, string.Concat(Common.Format((Common.ParseLarge(custom.Rectangle.Bottom.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration), "px")), context, configuration));
                            }
                            if (custom.Rectangle?.Left?.Value != null)
                            {
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.MarginLeftExact, string.Concat(Common.Format((Common.ParseLarge(custom.Rectangle.Left.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration), "px")), context, configuration));
                            }
                            if (custom.PathList != null)
                            {
                                (double X, double Y) last = (0, 0);
                                styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.ClippingAllPath, string.Join(' ', custom.PathList.Elements<DocumentFormat.OpenXml.Drawing.Path>().SelectMany(x => x.Elements()).Select(x =>
                                {
                                    switch (x)
                                    {
                                        case DocumentFormat.OpenXml.Drawing.CloseShapePath:
                                            return "Z";
                                        case DocumentFormat.OpenXml.Drawing.ArcTo arc:
                                            double width = (Common.ParseLarge(arc.WidthRadius?.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0;
                                            double height = (Common.ParseLarge(arc.HeightRadius?.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0;
                                            double start = (Common.ParseInteger(arc.StartAngle?.Value) * Common.RATIO_ANGLE * Math.PI / 180.0) ?? 0;
                                            double end = start + ((Common.ParseInteger(arc.SwingAngle?.Value) * Common.RATIO_ANGLE * Math.PI / 180.0) ?? 0);
                                            last = (last.X - width * Math.Cos(start) + width * Math.Cos(end), last.Y - height * Math.Sin(start) + height * Math.Sin(end));

                                            return Common.Format("A {0} {1} 0 1 1 {2},{3}", [Common.Format(width, configuration), Common.Format(height, configuration), Common.Format(last.X, configuration), Common.Format(last.Y, configuration)]);
                                        default:
                                            return string.Concat(x switch
                                            {
                                                DocumentFormat.OpenXml.Drawing.MoveTo => "M",
                                                DocumentFormat.OpenXml.Drawing.CubicBezierCurveTo => "C",
                                                DocumentFormat.OpenXml.Drawing.QuadraticBezierCurveTo => "Q",
                                                _ => "L",
                                            }, " ", string.Join(' ', x.Elements<DocumentFormat.OpenXml.Drawing.Point>().Select(y =>
                                            {
                                                last = ((Common.ParseLarge(y.X?.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, (Common.ParseLarge(y.Y?.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0);
                                                return string.Concat(Common.Format(last.X, configuration), ",", Common.Format(last.Y, configuration));
                                            })));
                                    }
                                }))), context, configuration));
                            }

                            break;
                    }
                }

                if ((isHidden ?? false) && configuration.ConvertVisibilities)
                {
                    element.Attributes[Common.ATTRIBUTE_HIDDEN] = null;
                }

                return element;
            }

            foreach (OpenXmlElement child in value.WorksheetDrawing.Elements())
            {
                (uint Index, string? Content) left = (0, null);
                (uint Index, string? Content) top = (0, null);
                (uint Index, string? Content) right = (0, null);
                (uint Index, string? Content) bottom = (0, null);
                switch (child)
                {
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absolute:
                        if (absolute.Position?.X?.Value != null)
                        {
                            double offset = absolute.Position.X.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                            left = (0, string.Concat(Common.Format(offset, configuration), "px"));
                            if (absolute.Extent?.Cx?.Value != null)
                            {
                                right = (0, Common.Format("calc(100% - {0}px)", [Common.Format(offset + absolute.Extent.Cx.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)]));
                            }
                        }
                        if (absolute.Position?.Y?.Value != null)
                        {
                            double offset = absolute.Position.Y.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                            top = (0, string.Concat(Common.Format(offset, configuration), "px"));
                            if (absolute.Extent?.Cy?.Value != null)
                            {
                                bottom = (0, Common.Format("calc(100% - {0}px)", [Common.Format(offset + absolute.Extent.Cy.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)]));
                            }
                        }

                        break;
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor single:
                        if (Common.ParsePositive(single.FromMarker?.ColumnId?.Text) is uint column)
                        {
                            double offset = (Common.ParseLarge(single.FromMarker?.ColumnOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0;
                            left = (column + 1, Common.Format("calc(var(--left) + {0}px)", [Common.Format(offset, configuration)]));
                            if (single.Extent?.Cx?.Value != null)
                            {
                                right = (0, Common.Format("calc(var(--left) - {0}px)", [Common.Format(offset + single.Extent.Cx.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)]));
                            }
                        }
                        if (Common.ParsePositive(single.FromMarker?.RowId?.Text) is uint row)
                        {
                            double offset = (Common.ParseLarge(single.FromMarker?.RowOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0;
                            top = (row + 1, Common.Format("calc(var(--top) + {0}px)", [Common.Format(offset, configuration)]));
                            if (single.Extent?.Cy?.Value != null)
                            {
                                bottom = (0, Common.Format("calc(var(--top) - {0}px)", [Common.Format(offset + single.Extent.Cy.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)]));
                            }
                        }

                        break;
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor dual:
                        if (Common.ParsePositive(dual.FromMarker?.ColumnId?.Text) is uint before)
                        {
                            left = (before + 1, Common.Format("calc(var(--left) + {0}px)", [Common.Format((Common.ParseLarge(dual.FromMarker?.ColumnOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)]));
                        }
                        if (Common.ParsePositive(dual.FromMarker?.RowId?.Text) is uint upper)
                        {
                            top = (upper + 1, Common.Format("calc(var(--top) + {0}px)", [Common.Format((Common.ParseLarge(dual.FromMarker?.RowOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)]));
                        }
                        if (Common.ParsePositive(dual.ToMarker?.ColumnId?.Text) is uint after)
                        {
                            right = (after + 1, Common.Format("calc(var(--right) - {0}px)", [Common.Format((Common.ParseLarge(dual.ToMarker?.ColumnOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)]));
                        }
                        if (Common.ParsePositive(dual.ToMarker?.RowId?.Text) is uint lower)
                        {
                            bottom = (lower + 1, Common.Format("calc(var(--bottom) - {0}px)", [Common.Format((Common.ParseLarge(dual.ToMarker?.RowOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)]));
                        }

                        break;
                }

                if (!(configuration.XlsxObjectSelector?.Invoke((left.Index > 0 && top.Index > 0 ? (left.Index, top.Index) : null, right.Index > 0 && bottom.Index > 0 ? (right.Index, bottom.Index) : null)) ?? true))
                {
                    continue;
                }

                Specification.Html.HtmlStyles positions = [.. configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.PositioningAbsolute), context, configuration)];
                if (top.Content != null)
                {
                    positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationTopExact, top.Content), context, configuration));
                }
                if (right.Content != null)
                {
                    positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationRightExact, right.Content), context, configuration));
                }
                if (bottom.Content != null)
                {
                    positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationBottomExact, bottom.Content), context, configuration));
                }
                if (left.Content != null)
                {
                    positions.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TranslationLeftExact, left.Content), context, configuration));
                }

                foreach (OpenXmlElement component in child.Elements())
                {
                    Specification.Html.HtmlStyles baseline = [.. configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BoundingInclusive), context, configuration)];
                    baseline.Apply(positions);

                    Specification.Html.HtmlElement? root = null;
                    switch (component)
                    {
                        case DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture when configuration.ConvertPictures:
                            Specification.Html.HtmlStyles dimension = [.. configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Image, Specification.Html.HtmlStyleType.WidthFull), context, configuration), .. configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Image, Specification.Html.HtmlStyleType.HeightFull), context, configuration)];
                            Specification.Html.HtmlElement image = new(Specification.Html.HtmlElementType.Unpaired, "img", new()
                            {
                                ["loading"] = "lazy",
                                ["decoding"] = "async",
                                [Common.ATTRIBUTE_STYLE] = dimension
                            });
                            root = new(Specification.Html.HtmlElementType.Paired, Common.TAG_CONTAINER, null, [image]);

                            if (picture.BlipFill?.Blip?.Embed?.Value != null && value.TryGetPartById(picture.BlipFill.Blip.Embed.Value, out OpenXmlPart? part) && part is ImagePart source)
                            {
                                using MemoryStream memory = new();
                                using Stream stream = source.GetStream();
                                stream.CopyTo(memory);
                                image.Attributes["src"] = Common.Format("data:{0};base64,{1}", [source.ContentType, System.Convert.ToBase64String(memory.ToArray())]);
                            }
                            if (picture.BlipFill?.SourceRectangle != null)
                            {
                                dimension.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Image, Specification.Html.HtmlStyleType.CroppingExact, Common.Format("inset({0}% {1}% {2}% {3}%)", [Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Top?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration), Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Right?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration), Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Bottom?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration), Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Left?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration)])), context, configuration));
                            }
                            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Title?.Value != null)
                            {
                                image.Attributes["title"] = WebUtility.HtmlEncode(picture.NonVisualPictureProperties.NonVisualDrawingProperties.Title.Value);
                            }
                            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value != null)
                            {
                                image.Attributes["alt"] = WebUtility.HtmlEncode(picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.Value);
                            }

                            //TODO: support for linked pictures

                            root = stylizer(root, baseline, picture.ShapeStyle, picture.ShapeProperties, picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Hidden?.Value);

                            break;
                        case DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape when configuration.ConvertShapes:
                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.PaddingAllExact, string.Concat(Common.Format(9.6, configuration), "px")), context, configuration));
                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextWrappingWrap), context, configuration));
                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.ClippingAllHidden), context, configuration));
                            Specification.Html.HtmlElement inner = new(Specification.Html.HtmlElementType.Paired, Common.TAG_CONTAINER);
                            root = inner;

                            foreach (OpenXmlElement body in shape.TextBody?.Elements() ?? [])
                            {
                                switch (body)
                                {
                                    case DocumentFormat.OpenXml.Drawing.Paragraph paragraph:
                                        Specification.Html.HtmlStyles individual = [];
                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginAllExact, "0"), context, configuration));
                                        Specification.Html.HtmlElement group = new(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT_GROUP, new()
                                        {
                                            [Common.ATTRIBUTE_STYLE] = individual
                                        });

                                        DocumentFormat.OpenXml.Drawing.TextCharacterPropertiesType? defaults = paragraph.GetFirstChild<DocumentFormat.OpenXml.Drawing.ParagraphProperties>()?.GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>();
                                        foreach (OpenXmlElement segment in paragraph.Elements())
                                        {
                                            switch (segment)
                                            {
                                                case DocumentFormat.OpenXml.Drawing.Break:
                                                    group.Children.Add(new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Unpaired, "br"));
                                                    break;
                                                case DocumentFormat.OpenXml.Drawing.Text text:
                                                    group.Children.Add(text.Text);
                                                    break;
                                                case DocumentFormat.OpenXml.Drawing.Run run when run.Text?.Text != null:
                                                    if (configuration.ConvertStyles)
                                                    {
                                                        Specification.Html.HtmlElement element = new(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, null, [run.Text.Text]);
                                                        Specification.Xlsx.XlsxStyles.ApplyStyles(element, [configuration.ConverterComposition.XlsxFontConverter.Convert(run.RunProperties ?? defaults, context, configuration)]);
                                                        group.Children.Add(element);
                                                    }
                                                    else
                                                    {
                                                        group.Children.Add(run.Text.Text);
                                                    }
                                                    break;
                                                case DocumentFormat.OpenXml.Drawing.ParagraphProperties properties:
                                                    if (properties.Alignment?.Value != null)
                                                    {
                                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, properties.Alignment.Value switch
                                                        {
                                                            _ when properties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left => Specification.Html.HtmlStyleType.AlignmentHorizontalLeft,
                                                            _ when properties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center => Specification.Html.HtmlStyleType.AlignmentHorizontalCenter,
                                                            _ when properties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right => Specification.Html.HtmlStyleType.AlignmentHorizontalRight,
                                                            _ => Specification.Html.HtmlStyleType.AlignmentHorizontalJustify,
                                                        }), context, configuration));
                                                    }
                                                    if (properties.LeftMargin?.Value != null)
                                                    {
                                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginLeftExact, string.Concat(Common.Format(properties.LeftMargin.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration), "px")), context, configuration));
                                                    }
                                                    if (properties.RightMargin?.Value != null)
                                                    {
                                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginRightExact, string.Concat(Common.Format(properties.RightMargin.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration), "px")), context, configuration));
                                                    }
                                                    if (properties.Indent?.Value != null)
                                                    {
                                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextIndentNumeric, Common.Format(properties.Indent.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)), context, configuration));
                                                    }
                                                    if (properties.DefaultTabSize?.Value != null)
                                                    {
                                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextTabNumeric, Common.Format(properties.DefaultTabSize.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)), context, configuration));
                                                    }
                                                    if (properties.RightToLeft?.Value != null)
                                                    {
                                                        individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.DirectionExact, properties.RightToLeft.Value ? "rtl" : "ltr"), context, configuration));
                                                    }

                                                    //TODO: support for text bullets

                                                    foreach (OpenXmlElement option in properties.Elements())
                                                    {
                                                        switch (option)
                                                        {
                                                            case DocumentFormat.OpenXml.Drawing.LineSpacing line when line.SpacingPoints?.Val?.Value != null:
                                                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextLineHeightNumeric, Common.Format(line.SpacingPoints.Val.Value * Common.RATIO_POINT_SPACING, configuration)), context, configuration));
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.LineSpacing line when line.SpacingPercent?.Val?.Value != null:
                                                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextLineHeightProportional, Common.Format(line.SpacingPercent.Val.Value * Common.RATIO_PERCENTAGE, configuration)), context, configuration));
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceBefore before when before.SpacingPercent?.Val?.Value != null:
                                                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginTopExact, string.Concat(Common.Format(before.SpacingPercent.Val.Value * Common.RATIO_PERCENTAGE, configuration), "em")), context, configuration));
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceBefore before when before.SpacingPoints?.Val?.Value != null:
                                                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginTopExact, string.Concat(Common.Format(before.SpacingPoints.Val.Value * Common.RATIO_POINT_SPACING, configuration), "px")), context, configuration));
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceAfter after when after.SpacingPercent?.Val?.Value != null:
                                                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginBottomExact, string.Concat(Common.Format(after.SpacingPercent.Val.Value * Common.RATIO_PERCENTAGE, configuration), "em")), context, configuration));
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceAfter after when after.SpacingPoints?.Val?.Value != null:
                                                                individual.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.MarginBottomExact, string.Concat(Common.Format(after.SpacingPoints.Val.Value * Common.RATIO_POINT_SPACING, configuration), "px")), context, configuration));
                                                                break;
                                                        }
                                                    }

                                                    break;
                                            }
                                        }

                                        inner.Children.Add(group);

                                        break;
                                    case DocumentFormat.OpenXml.Drawing.BodyProperties properties:
                                        if (properties.Anchor?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, properties.Anchor.Value switch
                                            {
                                                _ when properties.Anchor.Value == DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Center => Specification.Html.HtmlStyleType.AlignmentVerticalCenter,
                                                _ when properties.Anchor.Value == DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Bottom => Specification.Html.HtmlStyleType.AlignmentVerticalBottom,
                                                _ => Specification.Html.HtmlStyleType.AlignmentVerticalTop,
                                            }), context, configuration));
                                        }
                                        if (properties.Wrap?.Value != null && properties.Wrap.Value == DocumentFormat.OpenXml.Drawing.TextWrappingValues.None)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextWrappingNone), context, configuration));
                                        }
                                        if (properties.ColumnCount?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextColumnCountNumeric, Common.Format(properties.ColumnCount.Value, configuration)), context, configuration));
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.TextColumnGapNumeric, Common.Format((properties.ColumnSpacing?.Value * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)), context, configuration));
                                        }
                                        if (properties.RightToLeftColumns?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.DirectionExact, properties.RightToLeftColumns.Value ? "rtl" : "ltr"), context, configuration));
                                        }
                                        if (properties.HorizontalOverflow?.Value != null && properties.HorizontalOverflow.Value == DocumentFormat.OpenXml.Drawing.TextHorizontalOverflowValues.Overflow)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.ClippingHorizontalVisible), context, configuration));
                                        }
                                        if (properties.VerticalOverflow?.Value != null && properties.VerticalOverflow.Value == DocumentFormat.OpenXml.Drawing.TextVerticalOverflowValues.Overflow)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.ClippingVerticalVisible), context, configuration));
                                        }
                                        if (properties.TopInset?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.PaddingTopExact, string.Concat(properties.TopInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT, "px")), context, configuration));
                                        }
                                        if (properties.RightInset?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.PaddingRightExact, string.Concat(properties.RightInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT, "px")), context, configuration));
                                        }
                                        if (properties.BottomInset?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.PaddingBottomExact, string.Concat(properties.BottomInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT, "px")), context, configuration));
                                        }
                                        if (properties.LeftInset?.Value != null)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.PaddingLeftExact, string.Concat(properties.LeftInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT, "px")), context, configuration));
                                        }
                                        if (properties.Rotation?.Value != null && properties.Rotation.Value != 0)
                                        {
                                            inner = new(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, new()
                                            {
                                                [Common.ATTRIBUTE_STYLE] = new Specification.Html.HtmlStyles(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.RotationNumeric, Common.Format(properties.Rotation.Value * Common.RATIO_ANGLE, configuration)), context, configuration))
                                            }, root.Children);
                                            root.Children = [inner];
                                        }
                                        if (properties.Vertical?.Value != null && properties.Vertical.Value != DocumentFormat.OpenXml.Drawing.TextVerticalValues.Horizontal)
                                        {
                                            baseline.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, properties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.WordArtLeftToRight || properties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.MongolianVertical ? Specification.Html.HtmlStyleType.TextOrientationVerticalReverse : Specification.Html.HtmlStyleType.TextOrientationVertical), context, configuration));
                                            if (properties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.Vertical270)
                                            {
                                                inner = new(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, new()
                                                {
                                                    [Common.ATTRIBUTE_STYLE] = new Specification.Html.HtmlStyles(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Shape, Specification.Html.HtmlStyleType.RotationNumeric, "180"), context, configuration))
                                                }, root.Children);
                                                root.Children = [inner];
                                            }
                                        }

                                        break;
                                }
                            }

                            root = stylizer(root, baseline, shape.ShapeStyle, shape.ShapeProperties, shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value);

                            break;
                    }

                    if (root != null)
                    {
                        yield return new(root)
                        {
                            Range = new(left.Index, top.Index, right.Index, bottom.Index)
                        };
                    }
                }
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxColorConverter"/> class.
    /// </summary>
    public class DefaultXlsxColorConverter() : IConverterBase<OpenXmlElement?, string>
    {
        /// <summary>
        /// Represents the indexed colors.
        /// </summary>
        protected internal static (byte Red, byte Green, byte Blue)?[] INDICES =
        [
            (0, 0, 0),
            (255, 255, 255),
            (255, 0, 0),
            (0, 255, 0),
            (0, 0, 255),
            (255, 255, 0),
            (255, 0, 255),
            (0, 255, 255),
            (0, 0, 0),
            (255, 255, 255),
            (255, 0, 0),
            (0, 255, 0),
            (0, 0, 255),
            (255, 255, 0),
            (255, 0, 255),
            (0, 255, 255),
            (128, 0, 0),
            (0, 128, 0),
            (0, 0, 128),
            (128, 128, 0),
            (128, 0, 128),
            (0, 128, 128),
            (192, 192, 192),
            (128, 128, 128),
            (153, 153, 255),
            (153, 51, 102),
            (255, 255, 204),
            (204, 255, 255),
            (102, 0, 102),
            (255, 128, 128),
            (0, 102, 204),
            (204, 204, 255),
            (0, 0, 128),
            (255, 0, 255),
            (255, 255, 0),
            (0, 255, 255),
            (128, 0, 128),
            (128, 0, 0),
            (0, 128, 128),
            (0, 0, 255),
            (0, 204, 255),
            (204, 255, 255),
            (204, 255, 204),
            (255, 255, 153),
            (153, 204, 255),
            (255, 153, 204),
            (204, 153, 255),
            (255, 204, 153),
            (51, 102, 255),
            (51, 204, 204),
            (153, 204, 0),
            (255, 204, 0),
            (255, 153, 0),
            (255, 102, 0),
            (102, 102, 153),
            (150, 150, 150),
            (0, 51, 102),
            (51, 153, 102),
            (0, 51, 0),
            (51, 51, 0),
            (153, 51, 0),
            (153, 51, 102),
            (51, 51, 153),
            (51, 51, 51),
            (128, 128, 128),
            (255, 255, 255)
        ];

        /// <summary>
        /// Represents the system colors.
        /// </summary>
        protected internal static Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, (byte Red, byte Green, byte Blue)?> SYSTEMS = new()
        {
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveBorder] = (180, 180, 180),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ActiveCaption] = (153, 180, 209),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ApplicationWorkspace] = (171, 171, 171),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.Background] = (255, 255, 255),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonFace] = (240, 240, 240),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonHighlight] = (0, 120, 215),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonShadow] = (160, 160, 160),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ButtonText] = (0, 0, 0),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.CaptionText] = (0, 0, 0),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientActiveCaption] = (185, 209, 234),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.GradientInactiveCaption] = (215, 228, 242),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.GrayText] = (109, 109, 109),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.Highlight] = (0, 120, 215),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.HighlightText] = (255, 255, 255),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.HotLight] = (255, 165, 0),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveBorder] = (244, 247, 252),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaption] = (191, 205, 219),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.InactiveCaptionText] = (0, 0, 0),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoBack] = (255, 255, 225),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.InfoText] = (0, 0, 0),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.Menu] = (240, 240, 240),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuBar] = (240, 240, 240),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuHighlight] = (0, 120, 215),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.MenuText] = (0, 0, 0),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ScrollBar] = (200, 200, 200),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDDarkShadow] = (160, 160, 160),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.ThreeDLight] = (227, 227, 227),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.Window] = (255, 255, 255),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowFrame] = (100, 100, 100),
            [DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText] = (0, 0, 0)
        };

        /// <summary>
        /// Represents the preset colors.
        /// </summary>
        protected internal static Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, (byte Red, byte Green, byte Blue)?> PRESETS = new()
        {
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.AliceBlue] = (240, 248, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.AntiqueWhite] = (250, 235, 215),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Aqua] = (0, 255, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Aquamarine] = (127, 255, 212),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Azure] = (240, 255, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Beige] = (245, 245, 220),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Bisque] = (255, 228, 196),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Black] = (0, 0, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.BlanchedAlmond] = (255, 235, 205),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Blue] = (0, 0, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.BlueViolet] = (138, 43, 226),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Brown] = (165, 42, 42),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.BurlyWood] = (222, 184, 135),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.CadetBlue] = (95, 158, 160),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Chartreuse] = (127, 255, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Chocolate] = (210, 105, 30),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Coral] = (255, 127, 80),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.CornflowerBlue] = (100, 149, 237),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Cornsilk] = (255, 248, 220),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Crimson] = (220, 20, 60),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Cyan] = (0, 255, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue] = (0, 0, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan] = (0, 139, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod] = (184, 134, 11),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray] = (169, 169, 169),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen] = (0, 100, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki] = (189, 183, 107),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta] = (139, 0, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen] = (85, 107, 47),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange] = (255, 140, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid] = (153, 50, 204),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed] = (139, 0, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon] = (233, 150, 122),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen] = (143, 188, 143),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue] = (72, 61, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray] = (47, 79, 79),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise] = (0, 206, 209),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet] = (148, 0, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepPink] = (255, 20, 147),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DeepSkyBlue] = (0, 191, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGray] = (105, 105, 105),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DodgerBlue] = (30, 144, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Firebrick] = (178, 34, 34),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.FloralWhite] = (255, 250, 240),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.ForestGreen] = (34, 139, 34),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Fuchsia] = (255, 0, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Gainsboro] = (220, 220, 220),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.GhostWhite] = (248, 248, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Gold] = (255, 215, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Goldenrod] = (218, 165, 32),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Gray] = (128, 128, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Green] = (0, 128, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.GreenYellow] = (173, 255, 47),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Honeydew] = (240, 255, 240),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.HotPink] = (255, 105, 180),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.IndianRed] = (205, 92, 92),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Indigo] = (75, 0, 130),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Ivory] = (255, 255, 240),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Khaki] = (240, 230, 140),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Lavender] = (230, 230, 250),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LavenderBlush] = (255, 240, 245),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LawnGreen] = (124, 252, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LemonChiffon] = (255, 250, 205),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue] = (173, 216, 230),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral] = (240, 128, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan] = (224, 255, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow] = (250, 250, 210),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray] = (211, 211, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen] = (144, 238, 144),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink] = (255, 182, 193),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon] = (255, 160, 122),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen] = (32, 178, 170),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue] = (135, 206, 250),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray] = (119, 136, 153),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue] = (176, 196, 222),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow] = (255, 255, 224),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Lime] = (0, 255, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LimeGreen] = (50, 205, 50),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Linen] = (250, 240, 230),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Magenta] = (255, 0, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Maroon] = (128, 0, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MedAquamarine] = (102, 205, 170),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue] = (0, 0, 205),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid] = (186, 85, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple] = (147, 112, 219),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen] = (60, 179, 113),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue] = (123, 104, 238),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen] = (0, 250, 154),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise] = (72, 209, 204),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed] = (199, 21, 133),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MidnightBlue] = (25, 25, 112),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MintCream] = (245, 255, 250),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MistyRose] = (255, 228, 225),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Moccasin] = (255, 228, 181),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.NavajoWhite] = (255, 222, 173),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Navy] = (0, 0, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.OldLace] = (253, 245, 230),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Olive] = (128, 128, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.OliveDrab] = (107, 142, 35),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Orange] = (255, 165, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.OrangeRed] = (255, 69, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Orchid] = (218, 112, 214),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGoldenrod] = (238, 232, 170),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleGreen] = (152, 251, 152),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleTurquoise] = (175, 238, 238),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PaleVioletRed] = (219, 112, 147),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PapayaWhip] = (255, 239, 213),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PeachPuff] = (255, 218, 185),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Peru] = (205, 133, 63),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Pink] = (255, 192, 203),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Plum] = (221, 160, 221),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.PowderBlue] = (176, 224, 230),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Purple] = (128, 0, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Red] = (255, 0, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.RosyBrown] = (188, 143, 143),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.RoyalBlue] = (65, 105, 225),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SaddleBrown] = (139, 69, 19),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Salmon] = (250, 128, 114),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SandyBrown] = (244, 164, 96),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaGreen] = (46, 139, 87),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SeaShell] = (255, 245, 238),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Sienna] = (160, 82, 45),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Silver] = (192, 192, 192),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SkyBlue] = (135, 206, 235),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateBlue] = (106, 90, 205),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGray] = (112, 128, 144),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Snow] = (255, 250, 250),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SpringGreen] = (0, 255, 127),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SteelBlue] = (70, 130, 180),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Tan] = (210, 180, 140),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Teal] = (0, 128, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Thistle] = (216, 191, 216),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Tomato] = (255, 99, 71),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Turquoise] = (64, 224, 208),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Violet] = (238, 130, 238),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Wheat] = (245, 222, 179),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.White] = (255, 255, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.WhiteSmoke] = (245, 245, 245),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Yellow] = (255, 255, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.YellowGreen] = (154, 205, 50),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkBlue2010] = (0, 0, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkCyan2010] = (0, 139, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGoldenrod2010] = (184, 134, 11),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGray2010] = (169, 169, 169),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey2010] = (169, 169, 169),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGreen2010] = (0, 100, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkKhaki2010] = (189, 183, 107),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkMagenta2010] = (139, 0, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOliveGreen2010] = (85, 107, 47),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrange2010] = (255, 140, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkOrchid2010] = (153, 50, 204),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkRed2010] = (139, 0, 0),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSalmon2010] = (233, 150, 122),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSeaGreen2010] = (143, 188, 143),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateBlue2010] = (72, 61, 139),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGray2010] = (47, 79, 79),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey2010] = (47, 79, 79),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkTurquoise2010] = (0, 206, 209),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkViolet2010] = (148, 0, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightBlue2010] = (173, 216, 230),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCoral2010] = (240, 128, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightCyan2010] = (224, 255, 255),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGoldenrodYellow2010] = (250, 250, 210),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGray2010] = (211, 211, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey2010] = (211, 211, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGreen2010] = (144, 238, 144),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightPink2010] = (255, 182, 193),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSalmon2010] = (255, 160, 122),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSeaGreen2010] = (32, 178, 170),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSkyBlue2010] = (135, 206, 250),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGray2010] = (119, 136, 153),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey2010] = (119, 136, 153),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSteelBlue2010] = (176, 196, 222),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightYellow2010] = (255, 255, 224),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumAquamarine2010] = (102, 205, 170),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumBlue2010] = (0, 0, 205),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumOrchid2010] = (186, 85, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumPurple2010] = (147, 112, 219),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSeaGreen2010] = (60, 179, 113),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSlateBlue2010] = (123, 104, 238),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumSpringGreen2010] = (0, 250, 154),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumTurquoise2010] = (72, 209, 204),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.MediumVioletRed2010] = (199, 21, 133),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkGrey] = (169, 169, 169),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DimGrey] = (105, 105, 105),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.DarkSlateGrey] = (47, 79, 79),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.Grey] = (128, 128, 128),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightGrey] = (211, 211, 211),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.LightSlateGrey] = (119, 136, 153),
            [DocumentFormat.OpenXml.Drawing.PresetColorValues.SlateGrey] = (112, 128, 144)
        };

        /// <inheritdoc />
        public string Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return "currentColor";
            }

            double red = 0;
            double green = 0;
            double blue = 0;
            double alpha = 255;

            void parser(string hexadecimal)
            {
                hexadecimal = hexadecimal.TrimStart('#').PadLeft(8, 'F');
                alpha = Common.ParseHex(hexadecimal[..2]) ?? 255;
                red = Common.ParseHex(hexadecimal[2..4]) ?? 0;
                green = Common.ParseHex(hexadecimal[4..6]) ?? 0;
                blue = Common.ParseHex(hexadecimal[6..8]) ?? 0;
            }
            void modifier((Func<double, double>? Hue, Func<double, double>? Saturation, Func<double, double>? Luminance) transformers)
            {
                double[] rgb = [red / 255.0, green / 255.0, blue / 255.0];
                double maximum = rgb.Max();
                double minimum = rgb.Min();
                double chroma = maximum - minimum;
                double[] distances = maximum != minimum ? [.. rgb.Select(x => (maximum - x) / chroma)] : [0, 0, 0];

                double hue = maximum != minimum ? (60.0 * (maximum switch
                {
                    _ when maximum == rgb[0] => distances[2] - distances[1],
                    _ when maximum == rgb[1] => distances[0] - distances[2] + 2,
                    _ => distances[1] - distances[0] + 4
                }) % 360 + 360) % 360 : 0;
                if (transformers.Hue != null)
                {
                    hue = transformers.Hue(hue);
                }

                double saturation = maximum != minimum ? chroma / (1 - Math.Abs(maximum + minimum - 1)) : 0;
                if (transformers.Saturation != null)
                {
                    saturation = transformers.Saturation(saturation);
                }

                double luminance = (maximum + minimum) / 2;
                if (transformers.Luminance != null)
                {
                    luminance = transformers.Luminance(luminance);
                }
                if (saturation <= 0)
                {
                    red = Math.Clamp(255.0 * luminance, 0, 255);
                    green = Math.Clamp(255.0 * luminance, 0, 255);
                    blue = Math.Clamp(255.0 * luminance, 0, 255);
                    return;
                }

                double upper = luminance <= 0.5 ? luminance * (saturation + 1) : luminance + saturation - luminance * saturation;
                double lower = 2 * luminance - upper;

                double shifter(int index)
                {
                    double shifted = ((hue + 120.0 * (1 - index)) % 360 + 360) % 360;

                    return shifted switch
                    {
                        < 60 => lower + (upper - lower) * shifted / 60.0,
                        < 180 => upper,
                        < 240 => lower + (upper - lower) * (240 - shifted) / 60.0,
                        _ => lower
                    };
                }

                red = Math.Clamp(255.0 * shifter(0), 0, 255);
                green = Math.Clamp(255.0 * shifter(1), 0, 255);
                blue = Math.Clamp(255.0 * shifter(2), 0, 255);
            }
            bool aggregator(OpenXmlElement color, IEnumerable<OpenXmlElement> children)
            {
                switch (color)
                {
                    case DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgb when rgb.Val?.Value != null:
                        parser(rgb.Val.Value);
                        break;
                    case DocumentFormat.OpenXml.Drawing.RgbColorModelPercentage rgb:
                        red = Math.Clamp((255.0 * rgb.RedPortion?.Value * Common.RATIO_PERCENTAGE) ?? 0, 0, 255);
                        green = Math.Clamp((255.0 * rgb.GreenPortion?.Value * Common.RATIO_PERCENTAGE) ?? 0, 0, 255);
                        blue = Math.Clamp((255.0 * rgb.BluePortion?.Value * Common.RATIO_PERCENTAGE) ?? 0, 0, 255);
                        break;
                    case DocumentFormat.OpenXml.Drawing.HslColor hsl:
                        modifier((x => (hsl.HueValue?.Value * Common.RATIO_ANGLE) ?? 0, x => (hsl.SatValue?.Value * Common.RATIO_PERCENTAGE) ?? 0, x => (hsl.LumValue?.Value * Common.RATIO_PERCENTAGE) ?? 0));
                        break;
                    case DocumentFormat.OpenXml.Drawing.SystemColor key when key.Val?.Value != null && Common.Get(SYSTEMS, key.Val.Value) is (byte Red, byte Green, byte Blue) system:
                        red = system.Red;
                        green = system.Green;
                        blue = system.Blue;
                        break;
                    case DocumentFormat.OpenXml.Drawing.SystemColor fallback when fallback.LastColor?.Value != null:
                        parser(fallback.LastColor.Value);
                        break;
                    case DocumentFormat.OpenXml.Drawing.PresetColor key when key.Val?.Value != null && Common.Get(PRESETS, key.Val.Value) is (byte Red, byte Green, byte Blue) preset:
                        red = preset.Red;
                        green = preset.Green;
                        blue = preset.Blue;
                        break;
                    case DocumentFormat.OpenXml.Drawing.SchemeColor scheme when scheme.Val?.Value != null:
                        return ((DocumentFormat.OpenXml.Drawing.Color2Type?)(scheme.Val.Value switch
                        {
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Light1 => context.Theme?.ThemeElements?.ColorScheme?.Light1Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Dark1 => context.Theme?.ThemeElements?.ColorScheme?.Dark1Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Light2 => context.Theme?.ThemeElements?.ColorScheme?.Light2Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Dark2 => context.Theme?.ThemeElements?.ColorScheme?.Dark2Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 => context.Theme?.ThemeElements?.ColorScheme?.Accent1Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent2 => context.Theme?.ThemeElements?.ColorScheme?.Accent2Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent3 => context.Theme?.ThemeElements?.ColorScheme?.Accent3Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent4 => context.Theme?.ThemeElements?.ColorScheme?.Accent4Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent5 => context.Theme?.ThemeElements?.ColorScheme?.Accent5Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent6 => context.Theme?.ThemeElements?.ColorScheme?.Accent6Color,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Hyperlink => context.Theme?.ThemeElements?.ColorScheme?.Hyperlink,
                            _ when scheme.Val.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.FollowedHyperlink => context.Theme?.ThemeElements?.ColorScheme?.FollowedHyperlinkColor,
                            _ => null
                        }))?.FirstChild is OpenXmlElement child && aggregator(child, scheme.Elements());
                    default:
                        return false;
                }

                foreach (OpenXmlElement child in children)
                {
                    switch (child)
                    {
                        case DocumentFormat.OpenXml.Drawing.Shade shade when (shade.Val?.Value * Common.RATIO_PERCENTAGE) is double number:
                            red = Math.Clamp(red * number, 0, 255);
                            green = Math.Clamp(green * number, 0, 255);
                            blue = Math.Clamp(blue * number, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Tint tint when (tint.Val?.Value * Common.RATIO_PERCENTAGE) is double number:
                            red = Math.Clamp(red * number + 255.0 * (1 - number), 0, 255);
                            green = Math.Clamp(green * number + 255.0 * (1 - number), 0, 255);
                            blue = Math.Clamp(blue * number + 255.0 * (1 - number), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Inverse:
                            red = 255 - red;
                            green = 255 - green;
                            blue = 255 - blue;
                            break;
                        case DocumentFormat.OpenXml.Drawing.Gray:
                            double grayscale = red * 0.3 + green * 0.59 + blue * 0.11;
                            red = grayscale;
                            green = grayscale;
                            blue = grayscale;
                            break;
                        case DocumentFormat.OpenXml.Drawing.Complement:
                            double maximum = red > green ? (red > blue ? red : blue) : (green > blue ? green : blue);
                            red = maximum - red;
                            green = maximum - green;
                            blue = maximum - blue;
                            break;
                        case DocumentFormat.OpenXml.Drawing.Gamma:
                            red = Math.Clamp(255.0 * (red / 255.0 > 0.04045 ? Math.Pow((red / 255.0 + 0.055) / 1.055, 2.4) : red / 255.0 / 12.92), 0, 255);
                            green = Math.Clamp(255.0 * (green / 255.0 > 0.04045 ? Math.Pow((green / 255.0 + 0.055) / 1.055, 2.4) : green / 255.0 / 12.92), 0, 255);
                            blue = Math.Clamp(255.0 * (blue / 255.0 > 0.04045 ? Math.Pow((blue / 255.0 + 0.055) / 1.055, 2.4) : blue / 255.0 / 12.92), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.InverseGamma:
                            red = Math.Clamp(255.0 * (red / 255.0 > 0.0031308 ? 1.055 * Math.Pow(red / 255.0, 1 / 2.4) - 0.055 : red / 255.0 * 12.92), 0, 255);
                            green = Math.Clamp(255.0 * (green / 255.0 > 0.0031308 ? 1.055 * Math.Pow(green / 255.0, 1 / 2.4) - 0.055 : green / 255.0 * 12.92), 0, 255);
                            blue = Math.Clamp(255.0 * (blue / 255.0 > 0.0031308 ? 1.055 * Math.Pow(blue / 255.0, 1 / 2.4) - 0.055 : blue / 255.0 * 12.92), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Red channel when channel.Val?.Value != null:
                            red = Math.Clamp(255.0 * channel.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.RedModulation modulation when modulation.Val?.Value != null:
                            red = Math.Clamp(red * (modulation.Val.Value * Common.RATIO_PERCENTAGE), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.RedOffset offset when offset.Val?.Value != null:
                            red = Math.Clamp(red + 255.0 * offset.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Green channel when channel.Val?.Value != null:
                            green = Math.Clamp(255.0 * channel.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.GreenModulation modulation when modulation.Val?.Value != null:
                            green = Math.Clamp(green * (modulation.Val.Value * Common.RATIO_PERCENTAGE), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.GreenOffset offset when offset.Val?.Value != null:
                            green = Math.Clamp(green + 255.0 * offset.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Blue channel when channel.Val?.Value != null:
                            blue = Math.Clamp(255.0 * channel.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.BlueModulation modulation when modulation.Val?.Value != null:
                            blue = Math.Clamp(blue * (modulation.Val.Value * Common.RATIO_PERCENTAGE), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.BlueOffset offset when offset.Val?.Value != null:
                            blue = Math.Clamp(blue + 255.0 * offset.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Alpha channel when channel.Val?.Value != null:
                            alpha = Math.Clamp(255.0 * channel.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.AlphaModulation modulation when modulation.Val?.Value != null:
                            alpha = Math.Clamp(alpha * (modulation.Val.Value * Common.RATIO_PERCENTAGE), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.AlphaOffset offset when offset.Val?.Value != null:
                            alpha = Math.Clamp(alpha + 255.0 * offset.Val.Value * Common.RATIO_PERCENTAGE, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Hue channel when channel.Val?.Value != null:
                            modifier((x => channel.Val.Value * Common.RATIO_ANGLE, null, null));
                            break;
                        case DocumentFormat.OpenXml.Drawing.HueModulation modulation when modulation.Val?.Value != null:
                            modifier((x => x * (modulation.Val.Value * Common.RATIO_PERCENTAGE), null, null));
                            break;
                        case DocumentFormat.OpenXml.Drawing.HueOffset offset when offset.Val?.Value != null:
                            modifier((x => x + offset.Val.Value * Common.RATIO_ANGLE, null, null));
                            break;
                        case DocumentFormat.OpenXml.Drawing.Saturation channel when channel.Val?.Value != null:
                            modifier((null, x => channel.Val.Value * Common.RATIO_PERCENTAGE, null));
                            break;
                        case DocumentFormat.OpenXml.Drawing.SaturationModulation modulation when modulation.Val?.Value != null:
                            modifier((null, x => x * (modulation.Val.Value * Common.RATIO_PERCENTAGE), null));
                            break;
                        case DocumentFormat.OpenXml.Drawing.SaturationOffset offset when offset.Val?.Value != null:
                            modifier((null, x => x + offset.Val.Value * Common.RATIO_PERCENTAGE, null));
                            break;
                        case DocumentFormat.OpenXml.Drawing.Luminance channel when channel.Val?.Value != null:
                            modifier((null, null, x => channel.Val.Value * Common.RATIO_PERCENTAGE));
                            break;
                        case DocumentFormat.OpenXml.Drawing.LuminanceModulation modulation when modulation.Val?.Value != null:
                            modifier((null, null, x => x * (modulation.Val.Value * Common.RATIO_PERCENTAGE)));
                            break;
                        case DocumentFormat.OpenXml.Drawing.LuminanceOffset offset when offset.Val?.Value != null:
                            modifier((null, null, x => x + offset.Val.Value * Common.RATIO_PERCENTAGE));
                            break;
                        default:
                            break;
                    }
                }

                return true;
            }

            if (value is ColorType color)
            {
                if (color.Auto ?? false)
                {
                    return "initial";
                }
                else if (color.Rgb?.Value != null)
                {
                    parser(color.Rgb.Value);
                }
                else if (Common.Get(INDICES, color.Indexed?.Value) is (byte Red, byte Green, byte Blue) indexed)
                {
                    red = indexed.Red;
                    green = indexed.Green;
                    blue = indexed.Blue;
                }
                else if (color.Theme?.Value == null || ((DocumentFormat.OpenXml.Drawing.Color2Type?)(color.Theme.Value switch
                {
                    0 => context.Theme?.ThemeElements?.ColorScheme?.Light1Color,
                    1 => context.Theme?.ThemeElements?.ColorScheme?.Dark1Color,
                    2 => context.Theme?.ThemeElements?.ColorScheme?.Light2Color,
                    3 => context.Theme?.ThemeElements?.ColorScheme?.Dark2Color,
                    4 => context.Theme?.ThemeElements?.ColorScheme?.Accent1Color,
                    5 => context.Theme?.ThemeElements?.ColorScheme?.Accent2Color,
                    6 => context.Theme?.ThemeElements?.ColorScheme?.Accent3Color,
                    7 => context.Theme?.ThemeElements?.ColorScheme?.Accent4Color,
                    8 => context.Theme?.ThemeElements?.ColorScheme?.Accent5Color,
                    9 => context.Theme?.ThemeElements?.ColorScheme?.Accent6Color,
                    10 => context.Theme?.ThemeElements?.ColorScheme?.Hyperlink,
                    11 => context.Theme?.ThemeElements?.ColorScheme?.FollowedHyperlinkColor,
                    _ => null
                }))?.FirstChild is not OpenXmlElement child || !aggregator(child, child.Elements()))
                {
                    return "currentColor";
                }

                if (color.Tint?.Value != null && color.Tint.Value != 0)
                {
                    modifier((null, null, x => color.Tint.Value < 0 ? x * (1 + color.Tint.Value) : x * (1 - color.Tint.Value) + color.Tint.Value));
                }
            }
            else if (value.FirstChild is not OpenXmlElement child || !aggregator(child, child.Elements()))
            {
                return "currentColor";
            }

            int[] result = [(int)Math.Round(red), (int)Math.Round(green), (int)Math.Round(blue), (int)Math.Round(alpha)];

            return (configuration.UseHtmlHexadecimalColors, result[3] < 255) switch
            {
                (false, false) => Common.Format("rgb({0})", [string.Join(' ', result[..3].Select(x => Common.Format(x, configuration)))]),
                (false, true) => Common.Format("rgb({0} / {1})", [string.Join(' ', result[..3].Select(x => Common.Format(x, configuration))), Common.Format(result[3] / 255.0, configuration)]),
                (true, false) => string.Concat(["#", .. result[..3].Select(x => x.ToString("X2", CultureInfo.InvariantCulture))]),
                _ => string.Concat(["#", .. result.Select(x => x.ToString("X2", CultureInfo.InvariantCulture))]),
            };
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxStringConverter"/> class.
    /// </summary>
    public class DefaultXlsxStringConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxString>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxString Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxString result = new();
            StringBuilder builder = new();
            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case Text text:
                        builder.Append(text.Text);
                        result.Children.Add(text.Text);
                        break;
                    case Run run when run.Text?.Text != null:
                        builder.Append(run.Text.Text);
                        if (configuration.ConvertStyles)
                        {
                            Specification.Html.HtmlElement element = new(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, null, [run.Text.Text]);
                            Specification.Xlsx.XlsxStyles.ApplyStyles(element, [configuration.ConverterComposition.XlsxFontConverter.Convert(run.RunProperties, context, configuration)]);
                            result.Children.Add(element);
                        }
                        else
                        {
                            result.Children.Add(run.Text.Text);
                        }
                        break;
                }
            }

            result.Raw = builder.ToString();

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxFontConverter"/> class.
    /// </summary>
    public class DefaultXlsxFontConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxStylesLayer>
    {
        internal enum CommonStyleType
        {
            StrikethroughDouble,
            UnderlineHeavy,
            UnderlineDouble,
            UnderlineDashed,
            UnderlineDashedHeavy,
            UnderlineDotted,
            UnderlineDottedHeavy,
            UnderlineWavy,
            UnderlineWavyHeavy
        }

        /// <inheritdoc />
        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Html.HtmlAttributes stylizer(CommonStyleType type)
            {
                object key = (Common.CacheCategory.CommonStyles, type);
                if (Common.Get(context.Cache, key) is not Specification.Html.HtmlAttributes cache)
                {
                    Specification.Html.HtmlStyles styles = [];
                    styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextDecorationExact, type switch
                    {
                        CommonStyleType.StrikethroughDouble => "line-through double",
                        CommonStyleType.UnderlineDouble => "underline double",
                        CommonStyleType.UnderlineHeavy => "underline 4px",
                        CommonStyleType.UnderlineDashed => "underline dashed",
                        CommonStyleType.UnderlineDashedHeavy => "underline dashed 4px",
                        CommonStyleType.UnderlineDotted => "underline dotted",
                        CommonStyleType.UnderlineDottedHeavy => "underline dotted 4px",
                        CommonStyleType.UnderlineWavy => "underline wavy",
                        CommonStyleType.UnderlineWavyHeavy => "underline wavy 4px",
                        _ => "none"
                    }), context, configuration));

                    cache = new()
                    {
                        [Common.ATTRIBUTE_STYLE] = styles
                    };
                    context.Cache[key] = cache;
                }

                return cache;
            }

            Specification.Xlsx.XlsxStylesLayer result = new();
            List<string> decorations = [];

            if (value is DocumentFormat.OpenXml.Drawing.TextCharacterPropertiesType properties)
            {
                if (properties.FontSize?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextSizeNumeric, Common.Format(properties.FontSize.Value * Common.RATIO_POINT_SPACING, configuration)), context, configuration));
                }
                if (properties.Bold?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, properties.Bold.Value ? Specification.Html.HtmlStyleType.TextWeightBold : Specification.Html.HtmlStyleType.TextWeightNormal), context, configuration));
                }
                if (properties.Italic?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, properties.Italic.Value ? Specification.Html.HtmlStyleType.TextStyleItalic : Specification.Html.HtmlStyleType.TextStyleNormal), context, configuration));
                }
                if (properties.Strike?.Value != null)
                {
                    if (properties.Strike.Value == DocumentFormat.OpenXml.Drawing.TextStrikeValues.DoubleStrike)
                    {
                        result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, stylizer(CommonStyleType.StrikethroughDouble), x)]);
                    }
                    else if (properties.Strike.Value != DocumentFormat.OpenXml.Drawing.TextStrikeValues.NoStrike)
                    {
                        decorations.Add("line-through");
                    }
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextDecorationExact, "none"), context, configuration));
                }
                if (properties.Underline?.Value != null)
                {
                    if (properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single || properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Words)
                    {
                        decorations.Add("underline");
                    }
                    else if (properties.Underline.Value != DocumentFormat.OpenXml.Drawing.TextUnderlineValues.None)
                    {
                        result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, stylizer(properties.Underline.Value switch
                        {
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Double => CommonStyleType.UnderlineDouble,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dash => CommonStyleType.UnderlineDashed,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLong => CommonStyleType.UnderlineDashed,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDash => CommonStyleType.UnderlineDashed,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashHeavy => CommonStyleType.UnderlineDashedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLongHeavy => CommonStyleType.UnderlineDashedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDashHeavy => CommonStyleType.UnderlineDashedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dotted => CommonStyleType.UnderlineDotted,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDash => CommonStyleType.UnderlineDotted,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.HeavyDotted => CommonStyleType.UnderlineDottedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDashHeavy => CommonStyleType.UnderlineDottedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Wavy => CommonStyleType.UnderlineWavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyDouble => CommonStyleType.UnderlineWavyHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyHeavy => CommonStyleType.UnderlineWavyHeavy,
                            _ => CommonStyleType.UnderlineHeavy
                        }), x)]);
                    }
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextDecorationExact, "none"), context, configuration));
                }
                if (properties.Spacing?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextLetterSpacingNumeric, Common.Format(properties.Spacing.Value * Common.RATIO_POINT_SPACING, configuration)), context, configuration));
                }
                if (properties.Capital?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextTransformExact, properties.Capital.Value switch
                    {
                        _ when properties.Capital.Value == DocumentFormat.OpenXml.Drawing.TextCapsValues.All => "uppercase",
                        _ when properties.Capital.Value == DocumentFormat.OpenXml.Drawing.TextCapsValues.Small => "lowercase",
                        _ => "none"
                    }), context, configuration));
                }
            }

            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case Color foreground:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.ForegroundExact, configuration.ConverterComposition.XlsxColorConverter.Convert(foreground, context, configuration)), context, configuration));
                        break;
                    case FontSize size when size.Val?.Value != null:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextSizeNumeric, Common.Format(size.Val.Value * Common.RATIO_POINT, configuration)), context, configuration));
                        break;
                    case RunFont name when name.Val?.Value != null:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextFamilyTextual, WebUtility.HtmlEncode(name.Val.Value)), context, configuration));
                        break;
                    case FontName name when name.Val?.Value != null:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextFamilyTextual, WebUtility.HtmlEncode(name.Val.Value)), context, configuration));
                        break;
                    case Bold bold:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, (bold.Val?.Value ?? true) ? Specification.Html.HtmlStyleType.TextWeightBold : Specification.Html.HtmlStyleType.TextWeightNormal), context, configuration));
                        break;
                    case Italic italic:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, (italic.Val?.Value ?? true) ? Specification.Html.HtmlStyleType.TextStyleItalic : Specification.Html.HtmlStyleType.TextStyleNormal), context, configuration));
                        break;
                    case Strike strike:
                        if (strike.Val?.Value ?? true)
                        {
                            decorations.Add("line-through");
                        }
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextDecorationExact, "none"), context, configuration));
                        break;
                    case Underline underline:
                        if (underline.Val?.Value == UnderlineValues.Double)
                        {
                            result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, stylizer(CommonStyleType.UnderlineDouble), x)]);
                        }
                        else if (underline.Val?.Value != UnderlineValues.None)
                        {
                            decorations.Add("underline");
                        }
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextDecorationExact, "none"), context, configuration));
                        break;
                    case VerticalTextAlignment vertical when vertical.Val?.Value != null:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, vertical.Val.Value switch
                        {
                            _ when vertical.Val.Value == VerticalAlignmentRunValues.Superscript => Specification.Html.HtmlStyleType.AlignmentVerticalSuperscript,
                            _ when vertical.Val.Value == VerticalAlignmentRunValues.Subscript => Specification.Html.HtmlStyleType.AlignmentVerticalSubscript,
                            _ => Specification.Html.HtmlStyleType.AlignmentVerticalBaseline
                        }), context, configuration));
                        break;
                    case Extend extend:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, (extend.Val?.Value ?? true) ? Specification.Html.HtmlStyleType.TextStretchExpanded : Specification.Html.HtmlStyleType.TextStretchNormal), context, configuration));
                        break;
                    case Condense condense:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, (condense.Val?.Value ?? true) ? Specification.Html.HtmlStyleType.TextStretchCondensed : Specification.Html.HtmlStyleType.TextStretchNormal), context, configuration));
                        break;
                    case DocumentFormat.OpenXml.Drawing.NoFill:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.ForegroundNone), context, configuration));
                        break;
                    case DocumentFormat.OpenXml.Drawing.SolidFill foreground:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.ForegroundExact, configuration.ConverterComposition.XlsxColorConverter.Convert(foreground, context, configuration)), context, configuration));
                        break;
                    case DocumentFormat.OpenXml.Drawing.TextFontType name when name.Typeface?.Value != null:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextFamilyTextual, WebUtility.HtmlEncode(name.Typeface.Value)), context, configuration));
                        break;
                    case DocumentFormat.OpenXml.Drawing.Highlight highlight:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.BackgroundExact, configuration.ConverterComposition.XlsxColorConverter.Convert(highlight, context, configuration)), context, configuration));
                        break;
                    case DocumentFormat.OpenXml.Drawing.RightToLeft direction:
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.DirectionExact, (direction.Val?.Value ?? true) ? "rtl" : "ltr"), context, configuration));
                        break;
                }
            }

            if (decorations.Any())
            {
                result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(null, Specification.Html.HtmlStyleType.TextDecorationExact, string.Join(' ', decorations)), context, configuration));
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxFillConverter"/> class.
    /// </summary>
    public class DefaultXlsxFillConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxStylesLayer>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesLayer result = new();
            if (value is Fill fill)
            {
                if (fill.PatternFill != null && fill.PatternFill.PatternType?.Value != PatternValues.None)
                {
                    string foreground = configuration.ConverterComposition.XlsxColorConverter.Convert(fill.PatternFill.ForegroundColor != null ? fill.PatternFill.ForegroundColor : fill.PatternFill.BackgroundColor, context, configuration);
                    string? pattern = fill.PatternFill.PatternType?.Value switch
                    {
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkGray => "radial-gradient(circle at 1px 1px, {0} 0.5px, transparent 0) 0 0 / 3.2px 3.2px, radial-gradient(circle at 2.6px 2.6px, {0} 0.5px, transparent 0) 0 0 / 3.2px 3.2px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.MediumGray => "radial-gradient(circle at 1px 1px, {0} 0.5px, transparent 0) 0 0 / 3.6px 3.6px, radial-gradient(circle at 2.8px 2.8px, {0} 0.5px, transparent 0) 0 0 / 3.6px 3.6px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightGray => "radial-gradient(circle at 1px 1px, {0} 0.5px, transparent 0) 0 0 / 4px 4px, radial-gradient(circle at 3px 3px, {0} 0.5px, transparent 0) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.Gray125 => "radial-gradient(circle at 1px 1px, {0} 0.5px, transparent 0) 0 0 / 6px 6px, radial-gradient(circle at 4px 4px, {0} 0.5px, transparent 0) 0 0 / 6px 6px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.Gray0625 => "radial-gradient(circle at 1px 1px, {0} 0.5px, transparent 0) 0 0 / 9px 9px, radial-gradient(circle at 5.5px 5.5px, {0} 0.5px, transparent 0) 0 0 / 9px 9px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkHorizontal => "linear-gradient(0deg, {0} 1.5px, transparent 0) 0 0 / 100% 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightHorizontal => "linear-gradient(0deg, {0} 1px, transparent 0) 0 0 / 100% 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkVertical => "linear-gradient(90deg, {0} 1.5px, transparent 0) 0 0 / 4px 100%",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightVertical => "linear-gradient(90deg, {0} 1px, transparent 0) 0 0 / 4px 100%",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkDown => "linear-gradient(45deg, {0} 25%, transparent 25% 50%, {0} 50% 75%, transparent 75%) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightDown => "linear-gradient(45deg, {0} 10%, transparent 10% 50%, {0} 50% 60%, transparent 60%) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkUp => "linear-gradient(-45deg, {0} 25%, transparent 25% 50%, {0} 50% 75%, transparent 75%) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightUp => "linear-gradient(-45deg, {0} 10%, transparent 10% 50%, {0} 50% 60%, transparent 60%) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkGrid => "conic-gradient(transparent 90deg, {0} 90deg 180deg, transparent 180deg 270deg, {0} 270deg) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightGrid => "linear-gradient(90deg, {0} 1px, transparent 0) 0 0 / 4px 4px, linear-gradient(0deg, {0} 1px, transparent 0) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkTrellis => "linear-gradient(45deg, {0} 15%, transparent 15% 50%, {0} 50% 65%, transparent 65%) 0 0 / 4px 4px, linear-gradient(-45deg, {0} 15%, transparent 15% 50%, {0} 50% 65%, transparent 65%) 0 0 / 4px 4px",
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightTrellis => "linear-gradient(45deg, {0} 10%, transparent 10% 50%, {0} 50% 60%, transparent 60%) 0 0 / 4px 4px, linear-gradient(-45deg, {0} 10%, transparent 10% 50%, {0} 50% 60%, transparent 60%) 0 0 / 4px 4px",
                        _ => null
                    };
                    if (pattern != null)
                    {
                        pattern = Common.Format(pattern, [foreground]);
                        if (fill.PatternFill.BackgroundColor != null)
                        {
                            pattern = string.Concat(pattern, ", ", configuration.ConverterComposition.XlsxColorConverter.Convert(fill.PatternFill.BackgroundColor, context, configuration));
                        }
                    }
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.BackgroundExact, pattern ?? foreground), context, configuration));
                }
                else if (fill.GradientFill != null)
                {
                    if (fill.GradientFill.Type?.Value != GradientValues.Linear)
                    {
                        double left = fill.GradientFill.Left?.Value ?? 0;
                        double top = fill.GradientFill.Top?.Value ?? 0;
                        double right = fill.GradientFill.Right?.Value ?? 0;
                        double bottom = fill.GradientFill.Bottom?.Value ?? 0;
                        double radius = ((left + right) / 2 + (top + bottom) / 2 - left - top) / 2;
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.BackgroundExact, Common.Format("radial-gradient(circle at {0}% {1}%{2})", [Common.Format(100.0 * (left + right) / 2, configuration), Common.Format(100.0 * (top + bottom) / 2, configuration), string.Concat(fill.GradientFill.Elements<GradientStop>().Select(x => string.Concat(", ", configuration.ConverterComposition.XlsxColorConverter.Convert(x.Color, context, configuration), x.Position?.Value != null ? string.Concat(" ", Common.Format(100.0 * (radius + x.Position.Value * (1 - radius)), configuration), "%") : string.Empty)))])), context, configuration));
                    }
                    else
                    {
                        double degree = (((fill.GradientFill.Degree?.Value + 90) % 360 + 360) % 360) ?? 90;
                        result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.BackgroundExact, Common.Format("linear-gradient({0}deg{1})", [Common.Format(degree, configuration), string.Concat(fill.GradientFill.Elements<GradientStop>().Select(x => string.Concat(", ", configuration.ConverterComposition.XlsxColorConverter.Convert(x.Color, context, configuration), x.Position?.Value != null ? string.Concat(" ", Common.Format(100.0 * x.Position.Value, configuration), "%") : string.Empty)))])), context, configuration));
                    }
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxBorderConverter"/> class.
    /// </summary>
    public class DefaultXlsxBorderConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxStylesLayer>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesLayer result = new();

            void stylizer(BorderPropertiesType? border, Specification.Html.HtmlStyleType type)
            {
                if (border == null)
                {
                    return;
                }

                string? style = border.Style?.Value switch
                {
                    _ when border.Style?.Value == BorderStyleValues.Thick => "thick solid ",
                    _ when border.Style?.Value == BorderStyleValues.Medium => "medium solid ",
                    _ when border.Style?.Value == BorderStyleValues.MediumDashed => "medium dashed ",
                    _ when border.Style?.Value == BorderStyleValues.MediumDashDot => "medium dashed ",
                    _ when border.Style?.Value == BorderStyleValues.MediumDashDotDot => "medium dotted ",
                    _ when border.Style?.Value == BorderStyleValues.Double => "medium double ",
                    _ when border.Style?.Value == BorderStyleValues.Thin => "thin solid ",
                    _ when border.Style?.Value == BorderStyleValues.Dashed => "thin dashed ",
                    _ when border.Style?.Value == BorderStyleValues.DashDot => "thin dashed ",
                    _ when border.Style?.Value == BorderStyleValues.SlantDashDot => "thin dashed ",
                    _ when border.Style?.Value == BorderStyleValues.DashDotDot => "thin dotted ",
                    _ when border.Style?.Value == BorderStyleValues.Hair => "thin dotted ",
                    _ => null
                };
                if (style != null || border.Color != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, type, string.Concat(style, configuration.ConverterComposition.XlsxColorConverter.Convert(border.Color, context, configuration))), context, configuration));
                }
            }

            if (value is Border border)
            {
                stylizer(border.TopBorder, Specification.Html.HtmlStyleType.BorderTopExact);
                stylizer(border.RightBorder, Specification.Html.HtmlStyleType.BorderRightExact);
                stylizer(border.BottomBorder, Specification.Html.HtmlStyleType.BorderBottomExact);
                stylizer(border.LeftBorder, Specification.Html.HtmlStyleType.BorderLeftExact);
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxAlignmentConverter"/> class.
    /// </summary>
    public class DefaultXlsxAlignmentConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxStylesLayer>
    {
        /// <inheritdoc />
        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesLayer result = new();
            if (value is Alignment alignment)
            {
                if (alignment.Horizontal?.Value != null && alignment.Horizontal.Value != HorizontalAlignmentValues.General)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, alignment.Horizontal.Value switch
                    {
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.Left => Specification.Html.HtmlStyleType.AlignmentHorizontalLeft,
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.Center => Specification.Html.HtmlStyleType.AlignmentHorizontalCenter,
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.CenterContinuous => Specification.Html.HtmlStyleType.AlignmentHorizontalCenter,
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.Right => Specification.Html.HtmlStyleType.AlignmentHorizontalRight,
                        _ => Specification.Html.HtmlStyleType.AlignmentHorizontalJustify
                    }), context, configuration));
                }
                if (alignment.Vertical?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, alignment.Vertical.Value switch
                    {
                        _ when alignment.Vertical.Value == VerticalAlignmentValues.Top => Specification.Html.HtmlStyleType.AlignmentVerticalTop,
                        _ when alignment.Vertical.Value == VerticalAlignmentValues.Bottom => Specification.Html.HtmlStyleType.AlignmentVerticalBottom,
                        _ => Specification.Html.HtmlStyleType.AlignmentVerticalCenter
                    }), context, configuration));
                }
                if (alignment.Indent?.Value != null)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.TextIndentNumeric, Common.Format(alignment.Indent.Value, configuration)), context, configuration));
                }
                if (alignment.WrapText != null && (alignment.WrapText?.Value ?? true))
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.TextWrappingWrap), context, configuration));
                }
                if (alignment.TextRotation?.Value != null && alignment.TextRotation.Value != 0)
                {
                    Specification.Html.HtmlAttributes attributes = new()
                    {
                        [Common.ATTRIBUTE_STYLE] = new Specification.Html.HtmlStyles(configuration.ConverterComposition.HtmlStylizer.Convert(alignment.TextRotation.Value != 255 ? new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.RotationNumeric, alignment.TextRotation.Value > 90 ? Common.Format(alignment.TextRotation.Value - 90, configuration) : string.Concat("-", Common.Format(alignment.TextRotation.Value, configuration))) : new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.TextOrientationVertical), context, configuration))
                    };
                    result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, Common.TAG_TEXT, attributes, x)]);
                }
                if (alignment.ReadingOrder?.Value != null && alignment.ReadingOrder.Value != 0)
                {
                    result.Styles.Apply(configuration.ConverterComposition.HtmlStylizer.Convert(new(Specification.Html.HtmlStyleTarget.Cell, Specification.Html.HtmlStyleType.DirectionExact, alignment.ReadingOrder.Value > 1 ? "rtl" : "ltr"), context, configuration));
                }
            }

            return result;
        }
    }
}
