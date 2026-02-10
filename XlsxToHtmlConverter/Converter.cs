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

            T2 converter<T1, T2>(Base.IConverterBase<T1, T2> converter, T1 value)
            {
                return converter.Convert(value, context, configuration);
            }

            using StreamWriter writer = new(output, configuration.Encoding, configuration.BufferSize, true);
            int indent = 0;

            WorkbookPart? workbook = input.WorkbookPart;
            context.Theme = workbook?.ThemePart?.Theme;
            context.Stylesheet = converter(configuration.ConverterComposition.XlsxStylesheetReader, workbook?.WorkbookStylesPart?.Stylesheet);
            context.SharedStrings = converter(configuration.ConverterComposition.XlsxSharedStringTableReader, workbook?.SharedStringTablePart?.SharedStringTable);

            if (!configuration.UseHtmlFragment)
            {
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Declaration, "html")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "html")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "head")));
                indent++;

                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Unpaired, "meta", new()
                {
                    ["charset"] = configuration.Encoding.WebName
                })));

                if (configuration.HtmlTitle != null)
                {
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Paired, "title", null, [configuration.HtmlTitle])));
                }
            }

            writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Paired, "style", null, [configuration.HtmlPresetStylesheet])));
            if (configuration.ConvertStyles && configuration.UseHtmlClasses)
            {
                Base.Specification.Html.HtmlStylesCollection stylesheet = [];
                foreach (Base.Specification.Xlsx.XlsxBaseStyles styles in context.Stylesheet.BaseStyles)
                {
                    stylesheet[$".{styles.Name}"] = styles.GetsStyles();
                }
                foreach (Base.Specification.Xlsx.XlsxDifferentialStyles styles in context.Stylesheet.DifferentialStyles)
                {
                    stylesheet[$".{styles.Name}"] = styles.GetsStyles();
                }

                if (stylesheet.Any())
                {
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Paired, "style", null, [stylesheet])));
                }
            }

            if (!configuration.UseHtmlFragment)
            {
                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "head")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "body")));
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

                context.Sheet = converter(configuration.ConverterComposition.XlsxWorksheetReader, worksheet?.Worksheet);
                context.Sheet.Specialties.AddRange(worksheet?.TableDefinitionParts.SelectMany(x => converter(configuration.ConverterComposition.XlsxTableReader, x)) ?? []);
                context.Sheet.Specialties.AddRange(converter(configuration.ConverterComposition.XlsxDrawingReader, worksheet?.DrawingsPart));

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

                Base.Specification.Html.HtmlAttributes table = new()
                {
                    ["width"] = configuration.UseHtmlProportionalWidths ? "100%" : "fit-content"
                };
                if (sheet.State?.Value != null && sheet.State.Value != SheetStateValues.Visible && configuration.ConvertVisibilities)
                {
                    table["hidden"] = null;
                }

                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "table", table)));
                indent++;

                if (configuration.ConvertSheetTitles)
                {
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Paired, "caption", context.Sheet.TitleAttributes, sheet.Name?.Value != null ? [sheet.Name.Value] : null)));
                }

                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "colgroup")));
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
                        ["style"] = baseline
                    };
                    attributes.Merge(context.Sheet.ColumnAttributes);

                    if (configuration.ConvertSizes)
                    {
                        baseline["width"] = double.IsNaN(columns[i].Width) ? "auto" : $"{Base.Implementation.Common.Format(columns[i].Width, configuration)}{(configuration.UseHtmlProportionalWidths ? "%" : "ch")}";
                    }
                    if (columns[i].IsHidden != null && configuration.ConvertVisibilities)
                    {
                        baseline["visibility"] = (columns[i].IsHidden ?? false) ? "collapse" : "visible";
                    }

                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Unpaired, "col", attributes)));
                }

                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "colgroup")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "tbody")));
                indent++;

                (uint Column, uint Row) last = (context.Sheet.Dimension.ColumnStart - 1, context.Sheet.Dimension.RowStart - 1);
                List<Base.Specification.Xlsx.XlsxSpecialty> specialties = [];

                void content(uint column, uint row, Base.Specification.Html.HtmlElement? element = null)
                {
                    if (specialties.Any(x => x.Specialty is MergeCell && x.Range.ContainsColumn(column) && !x.Range.StartsAt(column, row) && context.Sheet.Dimension.Contains(x.Range.ColumnStart, x.Range.RowStart)))
                    {
                        return;
                    }

                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, element ?? new(indent, Base.Specification.Html.HtmlElementType.Paired, "td")));
                }
                void suffix()
                {
                    if (last.Row < context.Sheet.Dimension.RowStart)
                    {
                        return;
                    }

                    for (uint i = last.Column + 1; i <= context.Sheet.Dimension.ColumnEnd; i++)
                    {
                        content(i, last.Row);
                    }

                    indent--;
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "tr")));

                    callback?.Invoke(input, new(index, (last.Row - context.Sheet.Dimension.RowStart + 1, context.Sheet.Dimension.RowCount)));
                }

                Row? row = null;
                foreach (Base.Specification.Xlsx.XlsxCell entry in converter(configuration.ConverterComposition.XlsxWorksheetIterator, context.Sheet))
                {
                    Base.Specification.Xlsx.XlsxCell cell = entry;
                    if (!context.Sheet.Dimension.Contains(cell.Reference.Column, cell.Reference.Row))
                    {
                        continue;
                    }

                    while (cell.Reference.Row > last.Row)
                    {
                        suffix();

                        last = (context.Sheet.Dimension.ColumnStart - 1, last.Row + 1);
                        specialties = Base.Implementation.Common.Get(references, last.Row) ?? [];
                        row = cell.Reference.Row <= last.Row ? cell.Cell?.Parent as Row : null;

                        Base.Specification.Html.HtmlStyles baseline = [];
                        Base.Specification.Html.HtmlAttributes attributes = new()
                        {
                            ["style"] = baseline
                        };
                        attributes.Merge(context.Sheet.RowAttributes);

                        if (configuration.ConvertSizes)
                        {
                            baseline["height"] = $"{Base.Implementation.Common.Format((Base.Implementation.Common.Get(row?.Height?.Value, row != null ? row?.CustomHeight?.Value : false) * Base.Implementation.Common.RATIO_POINT) ?? context.Sheet.CellSize.Height, configuration)}px";
                        }
                        if (row?.Hidden?.Value != null && configuration.ConvertVisibilities)
                        {
                            baseline["visibility"] = row.Hidden.Value ? "collapse" : "visible";
                        }
                        if (tops.Contains(last.Row))
                        {
                            baseline["anchor-name"] = $"--row-{Base.Implementation.Common.Format(last.Row, configuration)}";
                        }

                        writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "tr", attributes)));
                        indent++;
                    }
                    for (uint i = last.Column + 1; i < cell.Reference.Column; i++)
                    {
                        content(i, cell.Reference.Row);
                    }

                    bool isSelected = configuration.XlsxCellSelector?.Invoke((cell.Reference.Column, cell.Reference.Row)) ?? true;

                    Base.Specification.Xlsx.XlsxBaseStyles? shared = Base.Implementation.Common.Get(context.Stylesheet.BaseStyles, cell.Cell?.StyleIndex?.Value ?? Base.Implementation.Common.Get(columns, cell.Reference.Column - context.Sheet.Dimension.ColumnStart).StylesIndex ?? row?.StyleIndex?.Value ?? 0);
                    if (shared != null)
                    {
                        cell.Styles.Add(shared);
                        cell.NumberFormat = shared.NumberFormatId != null ? Base.Implementation.Common.Get(context.Stylesheet.NumberFormats, shared.NumberFormatId.Value) : null;
                        cell.NumberFormatId = shared.NumberFormatId;
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
                                individual["overflow-x"] = "hidden";
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
                        cell.Attributes["class"] = new Base.Specification.Html.HtmlClasses();
                    }
                    cell.Attributes["style"] = individual;
                    cell.Attributes.Merge(context.Sheet.CellAttributes);

                    if (isSelected)
                    {
                        cell = converter(configuration.ConverterComposition.XlsxCellContentReader, cell);
                    }
                    Base.Specification.Html.HtmlElement element = new(indent, Base.Specification.Html.HtmlElementType.Paired, "td", cell.Attributes, cell.Children);

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
                        individual["content-visibility"] = "hidden";
                    }

                    content(cell.Reference.Column, cell.Reference.Row, element);
                    last = cell.Reference;
                }
                suffix();

                if (elements.Any())
                {
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "tr", new()
                    {
                        ["style"] = new Base.Specification.Html.HtmlStyles()
                        {
                            ["visibility"] = "collapse"
                        }
                    })));
                    indent++;

                    Base.Specification.Html.HtmlAttributes? anchor(uint column)
                    {
                        return lefts.Contains(column) ? new()
                        {
                            ["style"] = new Base.Specification.Html.HtmlStyles()
                            {
                                ["anchor-name"] = $"--column-{Base.Implementation.Common.Format(column, configuration)}"
                            }
                        } : null;
                    }

                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedStart, "td", anchor(context.Sheet.Dimension.ColumnStart))));
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
                            positions["--top"] = $"anchor(--row-{Base.Implementation.Common.Format(specialty.Range.RowStart, configuration)} top)";
                        }
                        if (specialty.Range.ColumnEnd > 0)
                        {
                            positions["--right"] = $"anchor(--column-{Base.Implementation.Common.Format(specialty.Range.ColumnEnd, configuration)} left)";
                        }
                        if (specialty.Range.RowEnd > 0)
                        {
                            positions["--bottom"] = $"anchor(--row-{Base.Implementation.Common.Format(specialty.Range.RowEnd, configuration)} top)";
                        }
                        if (specialty.Range.ColumnStart > 0)
                        {
                            positions["--left"] = $"anchor(--column-{Base.Implementation.Common.Format(specialty.Range.ColumnStart, configuration)} left)";
                        }
                        positions["visibility"] = "visible";

                        element.Indent = indent;
                        element.Attributes.MergeStyles(positions);

                        writer.Write(converter(configuration.ConverterComposition.HtmlWriter, element));
                    }

                    indent--;
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "td")));

                    for (uint i = context.Sheet.Dimension.ColumnStart + 1; i <= context.Sheet.Dimension.ColumnEnd; i++)
                    {
                        writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.Paired, "td", anchor(i))));
                    }

                    indent--;
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "tr")));
                }

                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "tbody")));

                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "table")));

                index = (index.Current + 1, index.Total);
            }

            indent--;
            if (!configuration.UseHtmlFragment)
            {
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "body")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.Specification.Html.HtmlElementType.PairedEnd, "html")));
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
        public const double RATIO_ANGLE = 1 / 60000.0;
        public const double RATIO_PERCENTAGE = 1 / 100000.0;
        public const double RATIO_POINT = 1 / 72.0 * 96.0;
        public const double RATIO_POINT_SPACING = 1 / 7200.0 * 96.0;
        public const double RATIO_ENGLISH_METRIC_UNIT = 1 / 914400.0 * 96.0;

        /// <summary>
        /// Specifies the category of a cached entry.
        /// </summary>
        public enum CacheCategory
        {
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
                double decimals => decimals.ToString(configuration.RoundingDigits < 0 ? "G" : $"0.{new string('#', configuration.RoundingDigits)}", CultureInfo.InvariantCulture),
                _ => value.ToString() ?? string.Empty
            };
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
        public string Convert(Specification.Html.HtmlElement value, ConverterContext context, ConverterConfiguration configuration)
        {
            string padding(int indent)
            {
                return string.Concat(Enumerable.Repeat(configuration.TabCharacter, indent));
            }
            string element(Specification.Html.HtmlElement element)
            {
                return element.Type switch
                {
                    Specification.Html.HtmlElementType.Declaration => $"<!DOCTYPE {element.Tag}>",
                    Specification.Html.HtmlElementType.Paired => $"<{element.Tag}{attributes(element.Attributes)}>{children(element.Children, element.Indent ?? 0)}</{element.Tag}>",
                    Specification.Html.HtmlElementType.PairedStart => $"<{element.Tag}{attributes(element.Attributes)}>",
                    Specification.Html.HtmlElementType.PairedEnd => $"</{element.Tag}>",
                    Specification.Html.HtmlElementType.Unpaired => $"<{element.Tag}{attributes(element.Attributes)}>",
                    _ => $"<!-- {children(element.Children, element.Indent ?? 0)} -->"
                };
            }
            string attributes(Specification.Html.HtmlAttributes attributes)
            {
                return string.Concat(attributes.Select(x => x.Value switch
                {
                    null => $" {x.Key}",
                    Specification.Html.HtmlClasses classes => classes.Any() ? $" {x.Key}=\"{string.Join(' ', classes)}\"" : string.Empty,
                    Specification.Html.HtmlStyles styles => styles.Any() ? $" {x.Key}=\"{string.Join(' ', styles.Select(y => $"{y.Key}: {y.Value};"))}\"" : string.Empty,
                    _ => $" {x.Key}=\"{x.Value}\""
                }));
            }
            string children(Specification.Html.HtmlChildren content, int indent)
            {
                return string.Concat(content.Select(x =>
                {
                    switch (x)
                    {
                        case Specification.Html.HtmlElement html:
                            return element(html);
                        case Specification.Html.HtmlStylesCollection css:
                            StringBuilder builder = new(configuration.NewlineCharacter);

                            foreach ((string selector, Specification.Html.HtmlStyles styles) in css)
                            {
                                builder.Append($"{padding(indent + 1)}{selector} {{{configuration.NewlineCharacter}");
                                foreach ((string property, string value) in styles)
                                {
                                    builder.Append($"{padding(indent + 2)}{property}: {value};{configuration.NewlineCharacter}");
                                }
                                builder.Append($"{padding(indent + 1)}}}{configuration.NewlineCharacter}");
                            }
                            builder.Append(padding(indent));

                            return builder.ToString();
                        default:
                            return raw(x.ToString());
                    }
                }));
            }
            string raw(string? raw)
            {
                if (raw == null)
                {
                    return string.Empty;
                }

                return WebUtility.HtmlEncode(raw);
            }

            return $"{padding(value.Indent ?? 0)}{element(value)}{configuration.NewlineCharacter}";
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxWorksheetIterator"/> class.
    /// </summary>
    public class DefaultXlsxWorksheetIterator() : IConverterBase<Specification.Xlsx.XlsxSheet?, IEnumerable<Specification.Xlsx.XlsxCell>>
    {
        public IEnumerable<Specification.Xlsx.XlsxCell> Convert(Specification.Xlsx.XlsxSheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                yield break;
            }

            (uint Column, uint Row) last = (value.Dimension.ColumnStart - 1, value.Dimension.RowStart - 1);

            (uint Column, uint Row) reference(string reference)
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
                    last = cell.CellReference?.Value != null ? reference(cell.CellReference.Value) : (index > last.Row ? value.Dimension.ColumnStart : last.Column + 1, index);

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
        public Specification.Xlsx.XlsxStylesCollection Convert(Stylesheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesCollection result = new();

            Specification.Xlsx.XlsxStylesLayer layer<T>(IConverterBase<T, Specification.Xlsx.XlsxStylesLayer> converter, T value)
            {
                return converter.Convert(value, context, configuration);
            }
            Specification.Xlsx.XlsxNumberFormat? codes(string? format)
            {
                if (WebUtility.HtmlDecode(format) is not string code || code.All(char.IsWhiteSpace))
                {
                    return null;
                }

                List<Specification.Xlsx.XlsxNumberFormatCode> codes = [new()];

                StringBuilder builder = new();
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
                                codes[^1]?.IsDate = true;
                                break;
                        }
                    }

                    builder.Append(character);
                }
                codes[^1]?.Code = builder.ToString();

                return codes.Count switch
                {
                    1 => new(codes[0]),
                    2 => new(codes[0], codes[1]),
                    3 => new(codes[0], codes[1], codes[2]),
                    _ => new(codes[0], codes[1], codes[2], codes[3])
                };
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
                    Name = configuration.UseHtmlClasses ? $"base-{Common.Format(i, configuration)}" : null,
                    IsHidden = Common.Get(cell.Protection?.Hidden?.Value, configuration.ConvertVisibilities ? cell.ApplyProtection?.Value : false)
                };

                if (configuration.ConvertStyles)
                {
                    styles.Styles.Merge(layer(configuration.ConverterComposition.XlsxFontConverter, Common.Get(fonts, cell.FontId?.Value, cell.ApplyFont?.Value)));
                    styles.Styles.Merge(layer(configuration.ConverterComposition.XlsxFillConverter, Common.Get(fills, cell.FillId?.Value, cell.ApplyFill?.Value)));
                    styles.Styles.Merge(layer(configuration.ConverterComposition.XlsxBorderConverter, Common.Get(borders, cell.BorderId?.Value, cell.ApplyBorder?.Value)));
                    styles.Styles.Merge(layer(configuration.ConverterComposition.XlsxAlignmentConverter, Common.Get(cell.Alignment, cell.ApplyAlignment?.Value)));
                }
                if (configuration.ConvertNumberFormats)
                {
                    styles.NumberFormatId = Common.Get(cell.NumberFormatId?.Value, cell.ApplyNumberFormat?.Value);
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
                    Name = configuration.UseHtmlClasses ? $"differential-{Common.Format(i, configuration)}" : null,
                    IsHidden = Common.Get(differential.Protection?.Hidden?.Value, configuration.ConvertVisibilities)
                };

                if (configuration.ConvertStyles)
                {
                    styles.FontStyles = layer(configuration.ConverterComposition.XlsxFontConverter, differential.Font);
                    styles.FillStyles = layer(configuration.ConverterComposition.XlsxFillConverter, differential.Fill);
                    styles.BorderStyles = layer(configuration.ConverterComposition.XlsxBorderConverter, differential.Border);
                    styles.AlignmentStyles = layer(configuration.ConverterComposition.XlsxAlignmentConverter, differential.Alignment);
                }
                if (configuration.ConvertNumberFormats)
                {
                    styles.NumberFormat = codes(differential.NumberingFormat?.FormatCode?.Value);
                }

                return styles;
            }) ?? []];

            foreach (NumberingFormat number in Common.Get(value.NumberingFormats, configuration.ConvertNumberFormats)?.Elements<NumberingFormat>() ?? [])
            {
                if (number.NumberFormatId?.Value == null || codes(number.FormatCode?.Value) is not Specification.Xlsx.XlsxNumberFormat format)
                {
                    continue;
                }

                result.NumberFormats[number.NumberFormatId.Value] = format;
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxSharedStringTableReader"/> class.
    /// </summary>
    public class DefaultXlsxSharedStringTableReader() : IConverterBase<SharedStringTable?, Specification.Xlsx.XlsxString[]>
    {
        public Specification.Xlsx.XlsxString[] Convert(SharedStringTable? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return [];
            }

            return [.. value.Elements().Select(x => configuration.ConverterComposition.XlsxStringConverter.Convert(x, context, configuration))];
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxWorksheetReader"/> class.
    /// </summary>
    public class DefaultXlsxWorksheetReader() : IConverterBase<Worksheet?, Specification.Xlsx.XlsxSheet>
    {
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
                        result.TitleAttributes["style"] = new Specification.Html.HtmlStyles()
                        {
                            ["--sheet-color"] = configuration.ConverterComposition.XlsxColorConverter.Convert(properties.TabColor, context, configuration)
                        };
                        break;
                    case SheetFormatProperties format:
                        if ((format.ZeroHeight?.Value ?? false) && configuration.ConvertVisibilities)
                        {
                            result.RowAttributes["style"] = new Specification.Html.HtmlStyles()
                            {
                                ["visibility"] = "collapse"
                            };
                        }

                        if (configuration.ConvertStyles)
                        {
                            Specification.Html.HtmlStyles styles = [];
                            if (format.ThickTop?.Value ?? false)
                            {
                                styles["border-top-width"] = "thick";
                            }
                            if (format.ThickBottom?.Value ?? false)
                            {
                                styles["border-bottom-width"] = "thick";
                            }

                            if (styles.Any())
                            {
                                result.CellAttributes["style"] = styles;
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
        internal enum CommonStyles
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

        internal static Dictionary<CommonStyles, Specification.Xlsx.XlsxBaseStyles> commons = new()
        {
            [CommonStyles.AlignmentCenter] = new()
            {
                Styles = new(new()
                {
                    ["text-align"] = "center"
                })
            },
            [CommonStyles.AlignmentRight] = new()
            {
                Styles = new(new()
                {
                    ["text-align"] = "right"
                })
            },
            [CommonStyles.AlignmentDistributed] = new()
            {
                Styles = new(new()
                {
                    ["display"] = "flex",
                    ["justify-content"] = "space-between"
                })
            },
            [CommonStyles.ColorBlack] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "black"
                })
            },
            [CommonStyles.ColorGreen] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "green"
                })
            },
            [CommonStyles.ColorWhite] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "white"
                })
            },
            [CommonStyles.ColorBlue] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "blue"
                })
            },
            [CommonStyles.ColorMagenta] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "magenta"
                })
            },
            [CommonStyles.ColorYellow] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "yellow"
                })
            },
            [CommonStyles.ColorCyan] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "cyan"
                })
            },
            [CommonStyles.ColorRed] = new()
            {
                Styles = new(new()
                {
                    ["color"] = "red"
                })
            }
        };
        internal static Dictionary<uint, Specification.Xlsx.XlsxNumberFormat> formats = new()
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
        internal static Dictionary<string, CommonStyles?> colors = new()
        {
            ["BLACK"] = CommonStyles.ColorBlack,
            ["GREEN"] = CommonStyles.ColorGreen,
            ["WHITE"] = CommonStyles.ColorWhite,
            ["BLUE"] = CommonStyles.ColorBlue,
            ["MAGENTA"] = CommonStyles.ColorMagenta,
            ["YELLOW"] = CommonStyles.ColorYellow,
            ["CYAN"] = CommonStyles.ColorCyan,
            ["RED"] = CommonStyles.ColorRed
        };
        internal static Dictionary<string, Func<double, double, bool>> conditions = new()
        {
            ["="] = (x, y) => x == y,
            ["<>"] = (x, y) => x != y,
            ["<"] = (x, y) => x < y,
            ["<="] = (x, y) => x <= y,
            [">"] = (x, y) => x > y,
            [">="] = (x, y) => x >= y
        };

        public Specification.Xlsx.XlsxCell Convert(Specification.Xlsx.XlsxCell? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new(null);
            }

            (string Raw, Specification.Html.HtmlChildren Children) text(Specification.Xlsx.XlsxString data)
            {
                return (data.Raw, data.Children);
            }
            Specification.Html.HtmlChildren data(object data, string raw)
            {
                if (!configuration.ConvertNumberFormats)
                {
                    return [raw];
                }

                //TODO: support for locale-dependent default formats

                (int section, Specification.Xlsx.XlsxNumberFormatCode? code) = (value.NumberFormat ?? (value.NumberFormatId != null ? Common.Get(formats, value.NumberFormatId.Value) : null)) is Specification.Xlsx.XlsxNumberFormat format ? data switch
                {
                    double number when number > 0 => (0, format.Positive),
                    double number when number < 0 => (1, format.Negative),
                    double number when number == 0 => (2, format.Zero),
                    _ => (3, format.Text)
                } : (-1, null);
                object? key = value.NumberFormatId != null ? (Common.CacheCategory.NumberFormat, value.NumberFormatId.Value, section) : null;

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
                        else if (Common.Get(colors, token) is CommonStyles color)
                        {
                            styles = Common.Get(commons, color, configuration.ConvertStyles);
                        }
                        else if (Common.Get(conditions, string.Concat(token.TakeWhile(x => x is '=' or '<' or '>'))) is Func<double, double, bool> comparator)
                        {
                            if (data is double number && Common.ParseDecimals(string.Concat(token.SkipWhile(x => x is '=' or '<' or '>'))) is double operand && comparator(number, operand))
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

                    code = new(code.Code[start..], code.IsDate);
                }

                if (code == null || code.Code.All(char.IsWhiteSpace) || code.Code.Trim().ToUpperInvariant() == "GENERAL")
                {
                    switch (data)
                    {
                        case DateTime date:
                            return time(date, [date.ToString("d", culture)]);
                        case double decimals:
                            string general = decimals.ToString(CultureInfo.InvariantCulture).Replace("+", string.Empty);
                            if (general.Length <= (general.StartsWith('-') ? 12 : 11))
                            {
                                return number(decimals, [general]);
                            }

                            string scientific = decimals.ToString("0.#######E0", CultureInfo.InvariantCulture);
                            return number(decimals, [decimals.ToString($"0.{new string('#', Math.Max(0, (scientific.StartsWith('-') ? 10 : 9) - (scientific.Length - scientific.IndexOf('E'))))}E0", CultureInfo.InvariantCulture)]);
                        default:
                            return [raw];
                    }
                }

                Specification.Html.HtmlChildren children = [];

                StringBuilder builder = new();
                if (code.IsDate)
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
                        information = tokens(code.Code, true, (x, y) => x switch
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
                        });

                        if (key != null)
                        {
                            context.Cache[key] = information;
                        }
                    }

                    bool time(int index)
                    {
                        (int Distance, bool? IsTime) left = (0, null);
                        (int Distance, bool? IsTime) right = (0, null);

                        for (int i = 1; index - i >= 0 && left.IsTime == null; i++)
                        {
                            left = information[index - i].FirstOrDefault(char.IsLetter) switch
                            {
                                'H' or 'S' => (left.Distance, true),
                                'Y' or 'D' => (left.Distance, false),
                                _ => (left.Distance + information[index - i].Length, null)
                            };
                        }
                        for (int i = 1; index + i < information.Count && right.IsTime == null && right.Distance <= left.Distance; i++)
                        {
                            right = information[index + i].FirstOrDefault(char.IsLetter) switch
                            {
                                'H' or 'S' => (right.Distance, true),
                                'Y' or 'D' => (right.Distance, false),
                                _ => (right.Distance + information[index + i].Length, null)
                            };
                        }

                        return (left.IsTime != right.IsTime && left.Distance > right.Distance ? right.IsTime : left.IsTime) ?? false;
                    }
                    TimeSpan duration()
                    {
                        if (date.Year < 100 || date.Year > 9999)
                        {
                            return TimeSpan.Zero;
                        }

                        return TimeSpan.FromDays(date.ToOADate());
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
                                suffix = $".{date.Millisecond.ToString(parts[^1], culture)}";
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
                            "M" => date.ToString(time(i) ? "%m" : "%M", culture),
                            "MM" => date.ToString(time(i) ? "mm" : "MM", culture),
                            "MMM" => date.ToString("MMM", culture),
                            "MMMM" => date.ToString("MMMM", culture),
                            "MMMMM" => new(date.ToString("MMMM", culture).FirstOrDefault(), 1),
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
                            "[H]" => duration().TotalHours.ToString("0", culture),
                            "[M]" => duration().TotalMinutes.ToString("0", culture),
                            "[MM]" => duration().TotalMinutes.ToString("00", culture),
                            "[S]" => duration().TotalSeconds.ToString("0", culture),
                            "[SS]" => duration().TotalSeconds.ToString("00", culture),
                            _ => null
                        } is string content)
                        {
                            builder.Append(content);
                        }
                        else
                        {
                            literal(builder, token);
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
                        information.Tokens = tokens(code.Code, false, (x, y) =>
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

                                    return y.Length > 0 && !(y[^1] is '0' or '#' or '?' or '/');
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
                            < 0.01 => (0, 1),
                            > 0.99 => (1, 1),
                            _ => null
                        };

                        (int Numerator, int Denominator) lower = (0, 1);
                        (int Numerator, int Denominator) upper = (1, 1);
                        while (fraction == null)
                        {
                            (int Numerator, int Denominator) middle = (lower.Numerator + upper.Numerator, lower.Denominator + upper.Denominator);
                            if (middle.Numerator < middle.Denominator * (remainder - 0.01))
                            {
                                lower = middle;
                            }
                            else if (middle.Numerator > middle.Denominator * (remainder + 0.01))
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
                    components[1] = information.Lengths[1] > 0 ? (number - integer).ToString($".{new string('#', information.Lengths[1])}", CultureInfo.InvariantCulture).TrimStart('.').PadRight(information.Lengths[1], ' ') : string.Empty;

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

                    stage = 0;
                    int index = 0;

                    int digit(string token, string source, int start)
                    {
                        int index = start;

                        foreach (char character in token)
                        {
                            builder.Append(source[index] is ' ' ? character switch
                            {
                                '0' => '0',
                                '?' => ' ',
                                _ => string.Empty
                            } : (stage < 1 && separators.Contains(index) ? $"{source[index]}{culture.NumberFormat.NumberGroupSeparator}" : source[index]));

                            index++;
                        }

                        return index;
                    }

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

                                    digit(left.PadLeft(numerator.Length, '0'), numerator.PadLeft(left.Length, ' '), 0);
                                    builder.Append('/');
                                    digit(right.PadRight(denominator.Length, '0'), denominator.PadRight(right.Length, ' '), 0);

                                    stage = 3;

                                    break;
                                }

                                if (stage != 1 && index <= 0)
                                {
                                    index = digit(new('0', components[stage].Length - information.Lengths[stage]), components[stage], index);
                                }

                                index = digit(token, components[stage], index);

                                break;
                            case '.' when stage < 1:
                                if (index <= 0)
                                {
                                    index = digit(new('0', components[0].Length - information.Lengths[0]), components[0], index);
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
                                builder.Append(sign is '-' || token.Length > 1 ? $"{token.First()}{sign}" : token.First());
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
                                literal(builder, token);
                                break;
                        }
                    }
                }
                children.Add(builder.ToString());

                if (children.Count > 1)
                {
                    Specification.Html.HtmlElement container = new(Specification.Html.HtmlElementType.Paired, "span", null, [.. children.Select(x => new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "span", null, [x]))]);
                    commons[CommonStyles.AlignmentDistributed].ApplyStyles(container);

                    children = [container];
                }

                return data switch
                {
                    DateTime date => time(date, children),
                    double decimals => number(decimals, children),
                    _ => [builder.ToString()]
                };
            }
            Specification.Html.HtmlChildren time(DateTime date, Specification.Html.HtmlChildren children)
            {
                if (!configuration.UseHtmlDataElements)
                {
                    return children;
                }

                return [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "time", new()
                {
                    ["datetime"] = date.ToString("yyyy-MM-ddThh:mm:ss.fff", CultureInfo.InvariantCulture)
                }, children)];
            }
            Specification.Html.HtmlChildren number(double decimals, Specification.Html.HtmlChildren children)
            {
                if (!configuration.UseHtmlDataElements)
                {
                    return children;
                }

                return [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "data", new()
                {
                    ["value"] = decimals.ToString(CultureInfo.InvariantCulture)
                }, children)];
            }
            List<string> tokens(string code, bool isStandardized, Func<char, StringBuilder, bool?> tokenizer, char[]? singles = null)
            {
                StringBuilder builder = new();
                List<string> tokens = [];

                bool isSpecial = false;
                foreach ((int index, char character, bool isEscaped) in Specification.Xlsx.XlsxNumberFormat.Escape(code, singles))
                {
                    if (isEscaped)
                    {
                        builder.Append(character);
                        continue;
                    }

                    char input = isStandardized ? char.ToUpperInvariant(character) : character;
                    bool? isAdditional = tokenizer(input, builder);

                    if (isAdditional ?? isSpecial)
                    {
                        tokens.Add(builder.ToString());
                        builder.Clear();
                    }

                    isSpecial = isAdditional != null;
                    builder.Append(isSpecial ? input : character);
                }
                tokens.Add(builder.ToString());

                return tokens;
            }
            void literal(StringBuilder builder, string content)
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

            string content = value.Cell?.CellValue?.Text ?? string.Empty;
            (string raw, Specification.Html.HtmlChildren? children) = value.Cell?.DataType?.Value switch
            {
                _ when value.Cell?.DataType?.Value == CellValues.Error => (content, [content]),
                _ when value.Cell?.DataType?.Value == CellValues.String => (content, [content]),
                _ when value.Cell?.DataType?.Value == CellValues.InlineString => text(configuration.ConverterComposition.XlsxStringConverter.Convert(value.Cell, context, configuration)),
                _ when value.Cell?.DataType?.Value == CellValues.SharedString => Common.ParsePositive(content) is uint index && Common.Get(context.SharedStrings, index) is Specification.Xlsx.XlsxString shared ? text(shared) : (string.Empty, []),
                _ when value.Cell?.DataType?.Value == CellValues.Boolean => (content, [content.Trim() switch {
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

                            bool equality(ConditionalFormattingOperatorValues operation)
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
                                _ when rule.Type.Value == ConditionalFormatValues.CellIs && rule.Operator?.Value != null => equality(rule.Operator.Value),
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
                    _ when value.Cell?.DataType?.Value == CellValues.Date => DateTime.TryParse(content, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date) ? (data(date, content), true) : ([content], false),
                    _ => Common.ParseDecimals(content) is double decimals ? (data(decimals, content), true) : (data(content, content), false)
                };

                if (isAligned && configuration.ConvertStyles)
                {
                    value.Styles.Insert(0, commons[CommonStyles.AlignmentRight]);
                }
            }
            else if ((value.Cell?.DataType?.Value == CellValues.Error || value.Cell?.DataType?.Value == CellValues.Boolean) && configuration.ConvertStyles)
            {
                value.Styles.Insert(0, commons[CommonStyles.AlignmentCenter]);
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
        internal enum TableComponent
        {
            Header,
            Data,
            Totals
        }

        public IEnumerable<Specification.Xlsx.XlsxSpecialty> Convert(TableDefinitionPart? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value?.Table == null || value.Table.Reference?.Value == null)
            {
                yield break;
            }

            Specification.Xlsx.XlsxRange range = new(value.Table.Reference.Value, context.Sheet.Dimension);
            uint start = value.Table.HeaderRowCount?.Value ?? 1;
            uint end = value.Table.TotalsRowCount?.Value ?? 0;
            uint middle = range.RowCount - start - end;

            Specification.Xlsx.XlsxDifferentialStyles? styles(uint? content, uint? border)
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

            if (start > 0 && styles(value.Table.HeaderRowFormatId?.Value, value.Table.HeaderRowBorderFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles header)
            {
                yield return new(header)
                {
                    Range = new(range.ColumnStart, range.RowStart, range.ColumnEnd, range.RowStart + start - 1)
                };
            }
            if (middle > 0 && styles(value.Table.DataFormatId?.Value, value.Table.BorderFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles data)
            {
                yield return new(data)
                {
                    Range = new(range.ColumnStart, range.RowStart + start, range.ColumnEnd, range.RowEnd - end)
                };
            }
            if (end > 0 && styles(value.Table.TotalsRowFormatId?.Value, value.Table.TotalsRowBorderFormatId?.Value) is Specification.Xlsx.XlsxDifferentialStyles totals)
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
        public IEnumerable<Specification.Xlsx.XlsxSpecialty> Convert(DrawingsPart? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                yield break;
            }

            string color(OpenXmlElement? color)
            {
                return configuration.ConverterComposition.XlsxColorConverter.Convert(color, context, configuration);
            }
            Specification.Html.HtmlElement styles(Specification.Html.HtmlElement element, Specification.Html.HtmlStyles styles, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeStyle? shape, DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties? properties, bool? isHidden)
            {
                element.Attributes["style"] = styles;

                if (shape != null)
                {
                    if (shape.FontReference != null)
                    {
                        styles["color"] = color(shape.FontReference);
                    }
                    if (shape.FillReference != null)
                    {
                        styles["background"] = color(shape.FillReference);
                    }
                    if (shape.LineReference != null)
                    {
                        styles["border"] = $"thin solid {color(shape.FontReference)}";
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
                        styles["filter"] = filter;
                    }
                }

                foreach (OpenXmlElement child in properties?.Elements() ?? [])
                {
                    switch (child)
                    {
                        case DocumentFormat.OpenXml.Drawing.Transform2D transform:
                            if (transform.Offset?.X?.Value != null && !styles.ContainsKey("left"))
                            {
                                double offset = transform.Offset.X.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                                styles["left"] = $"{Common.Format(offset, configuration)}px";

                                if (transform.Extents?.Cx?.Value != null && !styles.ContainsKey("right"))
                                {
                                    styles["right"] = $"calc(100% - {Common.Format(offset + transform.Extents.Cx.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px)";
                                }
                            }
                            if (transform.Offset?.Y?.Value != null && !styles.ContainsKey("top"))
                            {
                                double offset = transform.Offset.Y.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                                styles["top"] = $"{Common.Format(offset, configuration)}px";

                                if (transform.Extents?.Cy?.Value != null && !styles.ContainsKey("bottom"))
                                {
                                    styles["bottom"] = $"calc(100% - {Common.Format(offset + transform.Extents.Cy.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px)";
                                }
                            }
                            if (transform.Rotation?.Value != null)
                            {
                                styles["rotate"] = $"{Common.Format(transform.Rotation.Value * Common.RATIO_ANGLE, configuration)}deg";
                            }
                            if ((transform.HorizontalFlip?.Value ?? false) || (transform.VerticalFlip?.Value ?? false))
                            {
                                styles["scale"] = $"{((transform.HorizontalFlip?.Value ?? false) ? "-1" : "1")} {((transform.VerticalFlip?.Value ?? false) ? "-1" : "1")}";
                            }

                            break;
                        case DocumentFormat.OpenXml.Drawing.NoFill:
                            styles["background"] = "transparent";
                            break;
                        case DocumentFormat.OpenXml.Drawing.SolidFill background:
                            styles["background"] = color(background);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Outline outline:
                            if (outline.Width?.Value != null)
                            {
                                styles["border-width"] = $"{Common.Format(outline.Width.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px";
                            }
                            if (outline.CompoundLineType?.Value != null && outline.CompoundLineType.Value != DocumentFormat.OpenXml.Drawing.CompoundLineValues.Single)
                            {
                                styles["border-style"] = "double";
                            }

                            foreach (OpenXmlElement component in outline)
                            {
                                switch (component)
                                {
                                    case DocumentFormat.OpenXml.Drawing.PresetDash preset:
                                        styles["border-style"] = preset.Val?.Value switch
                                        {
                                            _ when preset.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid => "solid",
                                            _ when preset.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot => "dotted",
                                            _ when preset.Val?.Value == DocumentFormat.OpenXml.Drawing.PresetLineDashValues.SystemDashDotDot => "dotted",
                                            _ => "dashed",
                                        };
                                        break;
                                    case DocumentFormat.OpenXml.Drawing.CustomDash:
                                        styles["border-style"] = "dashed";
                                        break;
                                    case DocumentFormat.OpenXml.Drawing.NoFill:
                                        styles["background"] = "transparent";
                                        break;
                                    case DocumentFormat.OpenXml.Drawing.SolidFill border:
                                        styles["border-color"] = color(border);
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
                                styles["margin-top"] = $"{Common.Format((Common.ParseLarge(custom.Rectangle.Top.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px";
                            }
                            if (custom.Rectangle?.Right?.Value != null)
                            {
                                styles["margin-right"] = $"{Common.Format((Common.ParseLarge(custom.Rectangle.Right.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px";
                            }
                            if (custom.Rectangle?.Bottom?.Value != null)
                            {
                                styles["margin-bottom"] = $"{Common.Format((Common.ParseLarge(custom.Rectangle.Bottom.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px";
                            }
                            if (custom.Rectangle?.Left?.Value != null)
                            {
                                styles["margin-left"] = $"{Common.Format((Common.ParseLarge(custom.Rectangle.Left.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px";
                            }
                            if (custom.PathList != null)
                            {
                                (double X, double Y) last = (0, 0);
                                styles["clip-path"] = $"path('{string.Join(' ', custom.PathList.Elements<DocumentFormat.OpenXml.Drawing.Path>().SelectMany(x => x.Elements()).Select(x =>
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

                                            return $"A {Common.Format(width, configuration)} {Common.Format(height, configuration)} 0 1 1 {Common.Format(last.X, configuration)},{Common.Format(last.Y, configuration)}";
                                        default:
                                            return $"{x switch
                                            {
                                                DocumentFormat.OpenXml.Drawing.MoveTo => "M",
                                                DocumentFormat.OpenXml.Drawing.CubicBezierCurveTo => "C",
                                                DocumentFormat.OpenXml.Drawing.QuadraticBezierCurveTo => "Q",
                                                _ => "L",
                                            }} {string.Join(' ', x.Elements<DocumentFormat.OpenXml.Drawing.Point>().Select(y =>
                                            {
                                                last = ((Common.ParseLarge(y.X?.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, (Common.ParseLarge(y.Y?.Value) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0);
                                                return $"{Common.Format(last.X, configuration)},{Common.Format(last.Y, configuration)}";
                                            }))}";
                                    }
                                }))}')";
                            }

                            break;
                    }
                }

                if ((isHidden ?? false) && configuration.ConvertVisibilities)
                {
                    element.Attributes["hidden"] = null;
                }

                return element;
            }

            foreach (OpenXmlElement child in value.WorksheetDrawing.Elements())
            {
                Specification.Html.HtmlStyles positions = new()
                {
                    ["position"] = "absolute"
                };
                (uint Index, string? Field) left = (0, null);
                (uint Index, string? Field) top = (0, null);
                (uint Index, string? Field) right = (0, null);
                (uint Index, string? Field) bottom = (0, null);

                switch (child)
                {
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absolute:
                        if (absolute.Position?.X?.Value != null)
                        {
                            double offset = absolute.Position.X.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                            left = (0, $"{Common.Format(offset, configuration)}px");

                            if (absolute.Extent?.Cx?.Value != null)
                            {
                                right = (0, $"calc(100% - {Common.Format(offset + absolute.Extent.Cx.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px)");
                            }
                        }
                        if (absolute.Position?.Y?.Value != null)
                        {
                            double offset = absolute.Position.Y.Value * Common.RATIO_ENGLISH_METRIC_UNIT;
                            top = (0, $"{Common.Format(offset, configuration)}px");

                            if (absolute.Extent?.Cy?.Value != null)
                            {
                                bottom = (0, $"calc(100% - {Common.Format(offset + absolute.Extent.Cy.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px)");
                            }
                        }

                        break;
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor single:
                        if (Common.ParsePositive(single.FromMarker?.ColumnId?.Text) is uint column)
                        {
                            double offset = (Common.ParseLarge(single.FromMarker?.ColumnOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0;
                            left = (column + 1, $"calc(var(--left) + {Common.Format(offset, configuration)}px)");

                            if (single.Extent?.Cx?.Value != null)
                            {
                                right = (0, $"calc(var(--left) - {Common.Format(offset + single.Extent.Cx.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px)");
                            }
                        }
                        if (Common.ParsePositive(single.FromMarker?.RowId?.Text) is uint row)
                        {
                            double offset = (Common.ParseLarge(single.FromMarker?.RowOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0;
                            top = (row + 1, $"calc(var(--top) + {Common.Format(offset, configuration)}px)");

                            if (single.Extent?.Cy?.Value != null)
                            {
                                bottom = (0, $"calc(var(--top) - {Common.Format(offset + single.Extent.Cy.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px)");
                            }
                        }

                        break;
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor dual:
                        if (Common.ParsePositive(dual.FromMarker?.ColumnId?.Text) is uint before)
                        {
                            left = (before + 1, $"calc(var(--left) + {Common.Format((Common.ParseLarge(dual.FromMarker?.ColumnOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px)");
                        }
                        if (Common.ParsePositive(dual.FromMarker?.RowId?.Text) is uint upper)
                        {
                            top = (upper + 1, $"calc(var(--top) + {Common.Format((Common.ParseLarge(dual.FromMarker?.RowOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px)");
                        }
                        if (Common.ParsePositive(dual.ToMarker?.ColumnId?.Text) is uint after)
                        {
                            right = (after + 1, $"calc(var(--right) - {Common.Format((Common.ParseLarge(dual.ToMarker?.ColumnOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px)");
                        }
                        if (Common.ParsePositive(dual.ToMarker?.RowId?.Text) is uint lower)
                        {
                            bottom = (lower + 1, $"calc(var(--bottom) - {Common.Format((Common.ParseLarge(dual.ToMarker?.RowOffset?.Text) * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px)");
                        }

                        break;
                }

                if (top.Field != null)
                {
                    positions["top"] = top.Field;
                }
                if (right.Field != null)
                {
                    positions["right"] = right.Field;
                }
                if (bottom.Field != null)
                {
                    positions["bottom"] = bottom.Field;
                }
                if (left.Field != null)
                {
                    positions["left"] = left.Field;
                }

                foreach (OpenXmlElement component in child.Elements())
                {
                    Specification.Html.HtmlElement? root = null;
                    Specification.Html.HtmlStyles baseline = new()
                    {
                        ["box-sizing"] = "border-box"
                    };
                    baseline.Merge(positions);

                    switch (component)
                    {
                        case DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture when configuration.ConvertPictures:
                            Specification.Html.HtmlStyles dimension = new()
                            {
                                ["width"] = "100%",
                                ["height"] = "100%",
                            };
                            Specification.Html.HtmlElement image = new(Specification.Html.HtmlElementType.Unpaired, "img", new()
                            {
                                ["loading"] = "lazy",
                                ["decoding"] = "async",
                                ["style"] = dimension
                            });
                            root = new(Specification.Html.HtmlElementType.Paired, "div", null, [image]);

                            //TODO: support for linked pictures

                            if (picture.BlipFill?.Blip?.Embed?.Value != null && value.TryGetPartById(picture.BlipFill.Blip.Embed.Value, out OpenXmlPart? part) && part is ImagePart source)
                            {
                                using MemoryStream memory = new();
                                using Stream stream = source.GetStream();
                                stream.CopyTo(memory);

                                image.Attributes["src"] = $"data:{source.ContentType};base64,{System.Convert.ToBase64String(memory.ToArray())}";
                            }
                            if (picture.BlipFill?.SourceRectangle != null)
                            {
                                dimension["object-view-box"] = $"inset({Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Top?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration)}% {Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Right?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration)}% {Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Bottom?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration)}% {Common.Format((100.0 * picture.BlipFill?.SourceRectangle?.Left?.Value * Common.RATIO_PERCENTAGE) ?? 0, configuration)}%)";
                            }
                            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Title?.Value != null)
                            {
                                image.Attributes["title"] = WebUtility.HtmlEncode(picture.NonVisualPictureProperties.NonVisualDrawingProperties.Title.Value);
                            }
                            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value != null)
                            {
                                image.Attributes["alt"] = WebUtility.HtmlEncode(picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.Value);
                            }

                            root = styles(root, baseline, picture.ShapeStyle, picture.ShapeProperties, picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Hidden?.Value);

                            break;
                        case DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape when configuration.ConvertShapes:
                            baseline["padding"] = $"{Common.Format(9.6, configuration)}px";
                            baseline["white-space"] = "preserve wrap";
                            baseline["overflow"] = "clip";
                            Specification.Html.HtmlElement inner = new(Specification.Html.HtmlElementType.Paired, "div");
                            root = inner;

                            foreach (OpenXmlElement body in shape.TextBody?.Elements() ?? [])
                            {
                                switch (body)
                                {
                                    case DocumentFormat.OpenXml.Drawing.Paragraph paragraph:
                                        Specification.Html.HtmlStyles individual = new()
                                        {
                                            ["margin"] = "0"
                                        };
                                        Specification.Html.HtmlElement block = new(Specification.Html.HtmlElementType.Paired, "p", new()
                                        {
                                            ["style"] = individual
                                        });

                                        DocumentFormat.OpenXml.Drawing.TextCharacterPropertiesType? defaults = paragraph.GetFirstChild<DocumentFormat.OpenXml.Drawing.ParagraphProperties>()?.GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>();
                                        foreach (OpenXmlElement segment in paragraph.Elements())
                                        {
                                            switch (segment)
                                            {
                                                case DocumentFormat.OpenXml.Drawing.Break:
                                                    block.Children.Add(new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Unpaired, "br"));
                                                    break;
                                                case DocumentFormat.OpenXml.Drawing.Text text:
                                                    block.Children.Add(text.Text);
                                                    break;
                                                case DocumentFormat.OpenXml.Drawing.Run run when run.Text?.Text != null:
                                                    if (configuration.ConvertStyles)
                                                    {
                                                        Specification.Html.HtmlElement element = new(Specification.Html.HtmlElementType.Paired, "span", null, [run.Text.Text]);
                                                        Specification.Xlsx.XlsxStyles.ApplyStyles(element, [configuration.ConverterComposition.XlsxFontConverter.Convert(run.RunProperties ?? defaults, context, configuration)]);

                                                        block.Children.Add(element);
                                                    }
                                                    else
                                                    {
                                                        block.Children.Add(run.Text.Text);
                                                    }
                                                    break;
                                                case DocumentFormat.OpenXml.Drawing.ParagraphProperties properties:
                                                    if (properties.Alignment?.Value != null)
                                                    {
                                                        individual["text-align"] = properties.Alignment.Value switch
                                                        {
                                                            _ when properties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left => "left",
                                                            _ when properties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right => "right",
                                                            _ when properties.Alignment.Value == DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center => "center",
                                                            _ => "justify",
                                                        };
                                                    }
                                                    if (properties.LeftMargin?.Value != null)
                                                    {
                                                        individual["margin-left"] = $"{Common.Format(properties.LeftMargin.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px";
                                                    }
                                                    if (properties.RightMargin?.Value != null)
                                                    {
                                                        individual["margin-right"] = $"{Common.Format(properties.RightMargin.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px";
                                                    }
                                                    if (properties.Indent?.Value != null)
                                                    {
                                                        individual["text-indent"] = $"{Common.Format(properties.Indent.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px";
                                                    }
                                                    if (properties.DefaultTabSize?.Value != null)
                                                    {
                                                        individual["tab-size"] = $"{Common.Format(properties.DefaultTabSize.Value * Common.RATIO_ENGLISH_METRIC_UNIT, configuration)}px";
                                                    }
                                                    if (properties.RightToLeft?.Value != null)
                                                    {
                                                        individual["direction"] = properties.RightToLeft.Value ? "rtl" : "ltr";
                                                    }

                                                    //TODO: support for text bullets

                                                    foreach (OpenXmlElement option in properties.Elements())
                                                    {
                                                        switch (option)
                                                        {
                                                            case DocumentFormat.OpenXml.Drawing.LineSpacing line when line.SpacingPercent?.Val?.Value != null:
                                                                individual["line-height"] = Common.Format(line.SpacingPercent.Val.Value * Common.RATIO_PERCENTAGE, configuration);
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.LineSpacing line when line.SpacingPoints?.Val?.Value != null:
                                                                individual["line-height"] = $"{Common.Format(line.SpacingPoints.Val.Value * Common.RATIO_POINT_SPACING, configuration)}px";
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceBefore before when before.SpacingPercent?.Val?.Value != null:
                                                                individual["margin-top"] = $"{Common.Format(before.SpacingPercent.Val.Value * Common.RATIO_PERCENTAGE, configuration)}em";
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceBefore before when before.SpacingPoints?.Val?.Value != null:
                                                                individual["margin-top"] = $"{Common.Format(before.SpacingPoints.Val.Value * Common.RATIO_POINT_SPACING, configuration)}px";
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceAfter after when after.SpacingPercent?.Val?.Value != null:
                                                                individual["margin-bottom"] = $"{Common.Format(after.SpacingPercent.Val.Value * Common.RATIO_PERCENTAGE, configuration)}em";
                                                                break;
                                                            case DocumentFormat.OpenXml.Drawing.SpaceAfter after when after.SpacingPoints?.Val?.Value != null:
                                                                individual["margin-bottom"] = $"{Common.Format(after.SpacingPoints.Val.Value * Common.RATIO_POINT_SPACING, configuration)}px";
                                                                break;
                                                        }
                                                    }

                                                    break;
                                            }
                                        }

                                        inner.Children.Add(block);

                                        break;
                                    case DocumentFormat.OpenXml.Drawing.BodyProperties properties:
                                        if (properties.Anchor?.Value != null)
                                        {
                                            baseline["align-content"] = properties.Anchor.Value switch
                                            {
                                                _ when properties.Anchor.Value == DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Center => "center",
                                                _ when properties.Anchor.Value == DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Bottom => "end",
                                                _ => "start",
                                            };
                                        }
                                        if (properties.Wrap?.Value != null && properties.Wrap.Value == DocumentFormat.OpenXml.Drawing.TextWrappingValues.None)
                                        {
                                            baseline["white-space"] = "preserve nowrap";
                                        }
                                        if (properties.ColumnCount?.Value != null)
                                        {
                                            baseline["column-count"] = Common.Format(properties.ColumnCount.Value, configuration);
                                            baseline["column-gap"] = $"{Common.Format((properties.ColumnSpacing?.Value * Common.RATIO_ENGLISH_METRIC_UNIT) ?? 0, configuration)}px";
                                        }
                                        if (properties.RightToLeftColumns?.Value != null)
                                        {
                                            baseline["direction"] = properties.RightToLeftColumns.Value ? "rtl" : "ltr";
                                        }
                                        if (properties.HorizontalOverflow?.Value != null && properties.HorizontalOverflow.Value == DocumentFormat.OpenXml.Drawing.TextHorizontalOverflowValues.Overflow)
                                        {
                                            baseline["overflow-x"] = "visible";
                                        }
                                        if (properties.VerticalOverflow?.Value != null && properties.VerticalOverflow.Value == DocumentFormat.OpenXml.Drawing.TextVerticalOverflowValues.Overflow)
                                        {
                                            baseline["overflow-y"] = "visible";
                                        }
                                        if (properties.TopInset?.Value != null)
                                        {
                                            baseline["padding-top"] = $"{properties.TopInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT}px";
                                        }
                                        if (properties.RightInset?.Value != null)
                                        {
                                            baseline["padding-right"] = $"{properties.RightInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT}px";
                                        }
                                        if (properties.BottomInset?.Value != null)
                                        {
                                            baseline["padding-bottom"] = $"{properties.BottomInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT}px";
                                        }
                                        if (properties.LeftInset?.Value != null)
                                        {
                                            baseline["padding-left"] = $"{properties.LeftInset.Value * Common.RATIO_ENGLISH_METRIC_UNIT}px";
                                        }
                                        if (properties.Rotation?.Value != null && properties.Rotation.Value != 0)
                                        {
                                            inner = new(Specification.Html.HtmlElementType.Paired, "span", new()
                                            {
                                                ["style"] = new Specification.Html.HtmlStyles()
                                                {
                                                    ["display"] = "inline-block",
                                                    ["rotate"] = $"{Common.Format(properties.Rotation.Value * Common.RATIO_ANGLE, configuration)}deg"
                                                }
                                            }, root.Children);
                                            root.Children = [inner];
                                        }
                                        if (properties.Vertical?.Value != null && properties.Vertical.Value != DocumentFormat.OpenXml.Drawing.TextVerticalValues.Horizontal)
                                        {
                                            if (properties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.WordArtLeftToRight || properties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.MongolianVertical)
                                            {
                                                baseline["writing-mode"] = "vertical-lr";
                                                baseline["text-orientation"] = "upright";
                                            }
                                            else
                                            {
                                                baseline["writing-mode"] = "vertical-rl";
                                                baseline["text-orientation"] = "upright";
                                            }

                                            if (properties.Vertical.Value == DocumentFormat.OpenXml.Drawing.TextVerticalValues.Vertical270)
                                            {
                                                inner = new(Specification.Html.HtmlElementType.Paired, "span", new()
                                                {
                                                    ["style"] = new Specification.Html.HtmlStyles()
                                                    {
                                                        ["display"] = "inline-block",
                                                        ["rotate"] = "180deg"
                                                    }
                                                }, root.Children);
                                                root.Children = [inner];
                                            }
                                        }

                                        break;
                                }
                            }

                            root = styles(root, baseline, shape.ShapeStyle, shape.ShapeProperties, shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value);

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
        internal static (byte Red, byte Green, byte Blue)?[] indexes = [
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
            (255, 255, 255)];
        internal static Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, (byte Red, byte Green, byte Blue)?> systems = new()
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
        internal static Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, (byte Red, byte Green, byte Blue)?> presets = new()
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

            void hex(string hex)
            {
                hex = hex.TrimStart('#').PadLeft(8, 'F');
                alpha = Common.ParseHex(hex[..2]) ?? 255;
                red = Common.ParseHex(hex[2..4]) ?? 0;
                green = Common.ParseHex(hex[4..6]) ?? 0;
                blue = Common.ParseHex(hex[6..8]) ?? 0;
            }
            void modifier(Func<double, double> hue, Func<double, double> saturation, Func<double, double> luminance)
            {
                double[] rgb = [red / 255.0, green / 255.0, blue / 255.0];
                double maximum = rgb.Max();
                double minimum = rgb.Min();
                double chroma = maximum - minimum;
                double[] distances = maximum != minimum ? [.. rgb.Select(x => (maximum - x) / chroma)] : [0, 0, 0];

                double[] hsl = [hue(maximum != minimum ? (maximum switch
                {
                    _ when maximum == rgb[0] => distances[2] - distances[1],
                    _ when maximum == rgb[1] => distances[0] - distances[2] + 2,
                    _ => distances[1] - distances[0] + 4
                } * 60.0 % 360 + 360) % 360 : 0), saturation(maximum != minimum ? chroma / (1 - Math.Abs(maximum + minimum - 1)) : 0), luminance((maximum + minimum) / 2)];
                double upper = hsl[2] <= 0.5 ? hsl[2] * (hsl[1] + 1) : hsl[2] + hsl[1] - hsl[2] * hsl[1];
                double lower = 2 * hsl[2] - upper;

                for (int i = 0; i < 3; i++)
                {
                    if (hsl[1] <= 0)
                    {
                        rgb[i] = hsl[2];
                        continue;
                    }

                    double shifted = ((hsl[0] + (1 - i) * 120) % 360 + 360) % 360;
                    rgb[i] = shifted switch
                    {
                        < 60 => lower + (upper - lower) * shifted / 60.0,
                        < 180 => upper,
                        < 240 => lower + (upper - lower) * (240 - shifted) / 60.0,
                        _ => lower
                    };
                }

                red = Math.Clamp(255.0 * rgb[0], 0, 255);
                green = Math.Clamp(255.0 * rgb[1], 0, 255);
                blue = Math.Clamp(255.0 * rgb[2], 0, 255);
            }
            bool element(OpenXmlElement color, IEnumerable<OpenXmlElement> children)
            {
                switch (color)
                {
                    case DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgb when rgb.Val?.Value != null:
                        hex(rgb.Val.Value);
                        break;
                    case DocumentFormat.OpenXml.Drawing.RgbColorModelPercentage rgb:
                        red = Math.Clamp((255.0 * rgb.RedPortion?.Value * Common.RATIO_PERCENTAGE) ?? 0, 0, 255);
                        green = Math.Clamp((255.0 * rgb.GreenPortion?.Value * Common.RATIO_PERCENTAGE) ?? 0, 0, 255);
                        blue = Math.Clamp((255.0 * rgb.BluePortion?.Value * Common.RATIO_PERCENTAGE) ?? 0, 0, 255);
                        break;
                    case DocumentFormat.OpenXml.Drawing.HslColor hsl:
                        modifier(x => (hsl.HueValue?.Value * Common.RATIO_ANGLE) ?? 0, x => (hsl.SatValue?.Value * Common.RATIO_PERCENTAGE) ?? 0, x => (hsl.LumValue?.Value * Common.RATIO_PERCENTAGE) ?? 0);
                        break;
                    case DocumentFormat.OpenXml.Drawing.SystemColor key when key.Val?.Value != null && Common.Get(systems, key.Val.Value) is (byte Red, byte Green, byte Blue) system:
                        red = system.Red;
                        green = system.Green;
                        blue = system.Blue;
                        break;
                    case DocumentFormat.OpenXml.Drawing.SystemColor fallback when fallback.LastColor?.Value != null:
                        hex(fallback.LastColor.Value);
                        break;
                    case DocumentFormat.OpenXml.Drawing.PresetColor key when key.Val?.Value != null && Common.Get(presets, key.Val.Value) is (byte Red, byte Green, byte Blue) preset:
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
                        }))?.FirstChild is OpenXmlElement child && element(child, scheme.Elements());
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
                            double maximum = new[] { red, green, blue }.Max();
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
                            modifier(x => channel.Val.Value * Common.RATIO_ANGLE, x => x, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.HueModulation modulation when modulation.Val?.Value != null:
                            modifier(x => x * (modulation.Val.Value * Common.RATIO_PERCENTAGE), x => x, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.HueOffset offset when offset.Val?.Value != null:
                            modifier(x => x + offset.Val.Value * Common.RATIO_ANGLE, x => x, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Saturation channel when channel.Val?.Value != null:
                            modifier(x => x, x => channel.Val.Value * Common.RATIO_PERCENTAGE, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.SaturationModulation modulation when modulation.Val?.Value != null:
                            modifier(x => x, x => x * (modulation.Val.Value * Common.RATIO_PERCENTAGE), x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.SaturationOffset offset when offset.Val?.Value != null:
                            modifier(x => x, x => x + offset.Val.Value * Common.RATIO_PERCENTAGE, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Luminance channel when channel.Val?.Value != null:
                            modifier(x => x, x => x, x => channel.Val.Value * Common.RATIO_PERCENTAGE);
                            break;
                        case DocumentFormat.OpenXml.Drawing.LuminanceModulation modulation when modulation.Val?.Value != null:
                            modifier(x => x, x => x, x => x * (modulation.Val.Value * Common.RATIO_PERCENTAGE));
                            break;
                        case DocumentFormat.OpenXml.Drawing.LuminanceOffset offset when offset.Val?.Value != null:
                            modifier(x => x, x => x, x => x + offset.Val.Value * Common.RATIO_PERCENTAGE);
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
                    hex(color.Rgb.Value);
                }
                else if (Common.Get(indexes, color.Indexed?.Value) is (byte Red, byte Green, byte Blue) indexed)
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
                }))?.FirstChild is not OpenXmlElement child || !element(child, child.Elements()))
                {
                    return "currentColor";
                }

                if (color.Tint?.Value != null && color.Tint.Value != 0)
                {
                    modifier(x => x, x => x, x => color.Tint.Value < 0 ? x * (1 + color.Tint.Value) : x * (1 - color.Tint.Value) + color.Tint.Value);
                }
            }
            else if (value.FirstChild is not OpenXmlElement child || !element(child, child.Elements()))
            {
                return "currentColor";
            }

            int[] result = [(int)Math.Round(red), (int)Math.Round(green), (int)Math.Round(blue), (int)Math.Round(alpha)];

            return (configuration.UseHtmlHexColors, result[3] < 255) switch
            {
                (false, false) => $"rgb({string.Join(' ', result[..3].Select(x => Common.Format(x, configuration)))})",
                (false, true) => $"rgb({string.Join(' ', result[..3].Select(x => Common.Format(x, configuration)))} / {Common.Format(result[3] / 255.0, configuration)})",
                (true, false) => $"#{string.Concat(result[..3].Select(x => x.ToString("X2", CultureInfo.InvariantCulture)))}",
                _ => $"#{string.Concat(result.Select(x => x.ToString("X2", CultureInfo.InvariantCulture)))}",
            };
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxStringConverter"/> class.
    /// </summary>
    public class DefaultXlsxStringConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxString>
    {
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
                            Specification.Html.HtmlElement element = new(Specification.Html.HtmlElementType.Paired, "span", null, [run.Text.Text]);
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
        internal enum CommonStyles
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

        internal static Dictionary<CommonStyles, Specification.Html.HtmlAttributes> commons = new()
        {
            [CommonStyles.StrikethroughDouble] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "line-through double"
                }
            },
            [CommonStyles.UnderlineDouble] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline double"
                }
            },
            [CommonStyles.UnderlineHeavy] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline 4px"
                }
            },
            [CommonStyles.UnderlineDashed] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline dashed"
                }
            },
            [CommonStyles.UnderlineDashedHeavy] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline dashed 4px"
                }
            },
            [CommonStyles.UnderlineDotted] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline dotted"
                }
            },
            [CommonStyles.UnderlineDottedHeavy] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline dotted 4px"
                }
            },
            [CommonStyles.UnderlineWavy] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline wavy"
                }
            },
            [CommonStyles.UnderlineWavyHeavy] = new()
            {
                ["style"] = new Specification.Html.HtmlStyles()
                {
                    ["text-decoration"] = "underline wavy 4px"
                }
            }
        };

        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesLayer result = new();

            string color(OpenXmlElement? color)
            {
                return configuration.ConverterComposition.XlsxColorConverter.Convert(color, context, configuration);
            }

            List<string> decorations = [];

            if (value is DocumentFormat.OpenXml.Drawing.TextCharacterPropertiesType properties)
            {
                if (properties.FontSize?.Value != null)
                {
                    result.Styles["font-size"] = $"{Common.Format(properties.FontSize.Value * Common.RATIO_POINT_SPACING, configuration)}px";
                }
                if (properties.Bold?.Value != null)
                {
                    result.Styles["font-weight"] = properties.Bold.Value ? "bold" : "normal";
                }
                if (properties.Italic?.Value != null)
                {
                    result.Styles["font-style"] = properties.Italic.Value ? "italic" : "normal";
                }
                if (properties.Strike?.Value != null)
                {
                    if (properties.Strike.Value == DocumentFormat.OpenXml.Drawing.TextStrikeValues.DoubleStrike)
                    {
                        result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "span", commons[CommonStyles.StrikethroughDouble], x)]);
                    }
                    else if (properties.Strike.Value != DocumentFormat.OpenXml.Drawing.TextStrikeValues.NoStrike)
                    {
                        decorations.Add("line-through");
                    }
                    result.Styles["text-decoration"] = "none";
                }
                if (properties.Underline?.Value != null)
                {
                    if (properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single || properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Words)
                    {
                        decorations.Add("underline");
                    }
                    else if (properties.Underline.Value != DocumentFormat.OpenXml.Drawing.TextUnderlineValues.None)
                    {
                        result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "span", commons[properties.Underline.Value switch
                        {
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Double => CommonStyles.UnderlineDouble,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dash => CommonStyles.UnderlineDashed,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLong => CommonStyles.UnderlineDashed,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDash => CommonStyles.UnderlineDashed,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashHeavy => CommonStyles.UnderlineDashedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLongHeavy => CommonStyles.UnderlineDashedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDashHeavy => CommonStyles.UnderlineDashedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dotted => CommonStyles.UnderlineDotted,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDash => CommonStyles.UnderlineDotted,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.HeavyDotted => CommonStyles.UnderlineDottedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDashHeavy => CommonStyles.UnderlineDottedHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Wavy => CommonStyles.UnderlineWavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyDouble => CommonStyles.UnderlineWavyHeavy,
                            _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyHeavy => CommonStyles.UnderlineWavyHeavy,
                            _ => CommonStyles.UnderlineHeavy
                        }], x)]);
                    }
                    result.Styles["text-decoration"] = "none";
                }
                if (properties.Spacing?.Value != null)
                {
                    result.Styles["letter-spacing"] = $"{Common.Format(properties.Spacing.Value * Common.RATIO_POINT_SPACING, configuration)}px";
                }
                if (properties.Capital?.Value != null)
                {
                    result.Styles["text-transform"] = properties.Capital.Value switch
                    {
                        _ when properties.Capital.Value == DocumentFormat.OpenXml.Drawing.TextCapsValues.All => "uppercase",
                        _ when properties.Capital.Value == DocumentFormat.OpenXml.Drawing.TextCapsValues.Small => "lowercase",
                        _ => "none"
                    };
                }
            }

            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case Color foreground:
                        result.Styles["color"] = color(foreground);
                        break;
                    case FontSize size when size.Val?.Value != null:
                        result.Styles["font-size"] = $"{Common.Format(size.Val.Value * Common.RATIO_POINT, configuration)}px";
                        break;
                    case RunFont name when name.Val?.Value != null:
                        result.Styles["font-family"] = $"\'{WebUtility.HtmlEncode(name.Val.Value)}\'";
                        break;
                    case FontName name when name.Val?.Value != null:
                        result.Styles["font-family"] = $"\'{WebUtility.HtmlEncode(name.Val.Value)}\'";
                        break;
                    case Bold bold:
                        result.Styles["font-weight"] = (bold.Val?.Value ?? true) ? "bold" : "normal";
                        break;
                    case Italic italic:
                        result.Styles["font-style"] = (italic.Val?.Value ?? true) ? "italic" : "normal";
                        break;
                    case Strike strike:
                        if (strike.Val?.Value ?? true)
                        {
                            decorations.Add("line-through");
                        }
                        result.Styles["text-decoration"] = "none";
                        break;
                    case Underline underline:
                        if (underline.Val?.Value == UnderlineValues.Double)
                        {
                            result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "span", commons[CommonStyles.UnderlineDouble], x)]);
                        }
                        else if (underline.Val?.Value != UnderlineValues.None)
                        {
                            decorations.Add("underline");
                        }
                        result.Styles["text-decoration"] = "none";
                        break;
                    case VerticalTextAlignment vertical when vertical.Val?.Value != null:
                        result.Styles["vertical-align"] = vertical.Val.Value switch
                        {
                            _ when vertical.Val.Value == VerticalAlignmentRunValues.Subscript => "sub",
                            _ when vertical.Val.Value == VerticalAlignmentRunValues.Superscript => "super",
                            _ => "baseline"
                        };
                        break;
                    case Extend extend:
                        result.Styles["font-stretch"] = (extend.Val?.Value ?? true) ? "expanded" : "normal";
                        break;
                    case Condense condense:
                        result.Styles["font-stretch"] = (condense.Val?.Value ?? true) ? "condensed" : "normal";
                        break;
                    case DocumentFormat.OpenXml.Drawing.NoFill:
                        result.Styles["color"] = "transparent";
                        break;
                    case DocumentFormat.OpenXml.Drawing.SolidFill foreground:
                        result.Styles["color"] = color(foreground);
                        break;
                    case DocumentFormat.OpenXml.Drawing.TextFontType name when name.Typeface?.Value != null:
                        result.Styles["font-family"] = $"\'{WebUtility.HtmlEncode(name.Typeface.Value)}\'";
                        break;
                    case DocumentFormat.OpenXml.Drawing.Highlight highlight:
                        result.Styles["background-color"] = color(highlight);
                        break;
                    case DocumentFormat.OpenXml.Drawing.RightToLeft direction:
                        result.Styles["direction"] = (direction.Val?.Value ?? true) ? "rtl" : "ltr";
                        break;
                }
            }

            if (decorations.Any())
            {
                result.Styles["text-decoration"] = string.Join(' ', decorations);
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxFillConverter"/> class.
    /// </summary>
    public class DefaultXlsxFillConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxStylesLayer>
    {
        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesLayer result = new();

            string color(OpenXmlElement? color)
            {
                return configuration.ConverterComposition.XlsxColorConverter.Convert(color, context, configuration);
            }

            if (value is Fill fill)
            {
                if (fill.PatternFill != null && fill.PatternFill.PatternType?.Value != PatternValues.None)
                {
                    if (fill.PatternFill.BackgroundColor != null)
                    {
                        result.Styles["background"] = color(fill.PatternFill.BackgroundColor);
                    }

                    string foreground = color(fill.PatternFill.ForegroundColor);
                    (string? image, string? size, string? position) = fill.PatternFill.PatternType?.Value switch
                    {
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkGray => ($"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 2.6px 2.6px, {foreground} 0.5px, transparent 0)", "3.2px 3.2px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.MediumGray => ($"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 2.8px 2.8px, {foreground} 0.5px, transparent 0)", "3.6px 3.6px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightGray => ($"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 3px 3px, {foreground} 0.5px, transparent 0)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.Gray125 => ($"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 4px 4px, {foreground} 0.5px, transparent 0)", "6px 6px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.Gray0625 => ($"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 5.5px 5.5px, {foreground} 0.5px, transparent 0)", "9px 9px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkHorizontal => ($"linear-gradient(0deg, {foreground} 1.5px, transparent 0)", "100% 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightHorizontal => ($"linear-gradient(0deg, {foreground} 1px, transparent 0)", "100% 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkVertical => ($"linear-gradient(90deg, {foreground} 1.5px, transparent 0)", "4px 100%", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightVertical => ($"linear-gradient(90deg, {foreground} 1px, transparent 0)", "4px 100%", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkDown => ($"linear-gradient(45deg, {foreground} 25%, transparent 25% 50%, {foreground} 50% 75%, transparent 75%)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightDown => ($"linear-gradient(45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkUp => ($"linear-gradient(-45deg, {foreground} 25%, transparent 25% 50%, {foreground} 50% 75%, transparent 75%)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightUp => ($"linear-gradient(-45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkGrid => ($"linear-gradient(45deg, {foreground} 25%, transparent 25% 75%, {foreground} 75%), linear-gradient(45deg, {foreground} 25%, transparent 25% 75%, {foreground} 75%)", "0 0, 2.5px 2.5px", "5px 5px"),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightGrid => ($"linear-gradient(90deg, {foreground} 1px, transparent 0), linear-gradient(0deg, {foreground} 1px, transparent 0)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkTrellis => ($"linear-gradient(45deg, {foreground} 15%, transparent 15% 50%, {foreground} 50% 65%, transparent 65%), linear-gradient(-45deg, {foreground} 15%, transparent 15% 50%, {foreground} 50% 65%, transparent 65%)", "4px 4px", null),
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightTrellis => ($"linear-gradient(45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%), linear-gradient(-45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%)", "4px 4px", null),
                        _ => (null, null, null)
                    };

                    if (image != null)
                    {
                        result.Styles["background-image"] = image;
                    }
                    else if (fill.PatternFill.ForegroundColor != null)
                    {
                        result.Styles["background"] = foreground;
                    }

                    if (size != null)
                    {
                        result.Styles["background-size"] = size;
                    }
                    if (position != null)
                    {
                        result.Styles["background-position"] = position;
                    }
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

                        IEnumerable<GradientStop> stops = fill.GradientFill.Elements<GradientStop>();
                        result.Styles["background"] = $"radial-gradient(circle at {Common.Format(100.0 * (left + right) / 2, configuration)}% {Common.Format(100.0 * (top + bottom) / 2, configuration)}%{string.Concat(stops.Select(x => $", {color(x.Color)}{(x.Position?.Value != null ? $" {Common.Format(100.0 * (radius + x.Position.Value * (1 - radius)), configuration)}%" : string.Empty)}"))})";
                    }
                    else
                    {
                        double degree = (((fill.GradientFill.Degree?.Value + 90) % 360 + 360) % 360) ?? 90;

                        IEnumerable<GradientStop> stops = fill.GradientFill.Elements<GradientStop>();
                        result.Styles["background"] = $"linear-gradient({Common.Format(degree, configuration)}deg{string.Concat(stops.Select(x => $", {color(x.Color)}{(x.Position?.Value != null ? $" {Common.Format(100.0 * x.Position.Value, configuration)}%" : string.Empty)}"))})";
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
        public Specification.Xlsx.XlsxStylesLayer Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            Specification.Xlsx.XlsxStylesLayer result = new();

            string? styles(BorderPropertiesType? border)
            {
                if (border == null)
                {
                    return null;
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

                return style != null || border.Color != null ? $"{style}{configuration.ConverterComposition.XlsxColorConverter.Convert(border.Color, context, configuration)}" : null;
            }

            if (value is Border border)
            {
                if (styles(border.TopBorder) is string top)
                {
                    result.Styles["border-top"] = top;
                }
                if (styles(border.RightBorder) is string right)
                {
                    result.Styles["border-right"] = right;
                }
                if (styles(border.BottomBorder) is string bottom)
                {
                    result.Styles["border-bottom"] = bottom;
                }
                if (styles(border.LeftBorder) is string left)
                {
                    result.Styles["border-left"] = left;
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxAlignmentConverter"/> class.
    /// </summary>
    public class DefaultXlsxAlignmentConverter : IConverterBase<OpenXmlElement?, Specification.Xlsx.XlsxStylesLayer>
    {
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
                    result.Styles["text-align"] = alignment.Horizontal.Value switch
                    {
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.Left => "left",
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.Right => "right",
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.Center => "center",
                        _ when alignment.Horizontal.Value == HorizontalAlignmentValues.CenterContinuous => "center",
                        _ => "justify"
                    };
                }
                if (alignment.Vertical?.Value != null)
                {
                    result.Styles["vertical-align"] = alignment.Vertical.Value switch
                    {
                        _ when alignment.Vertical.Value == VerticalAlignmentValues.Bottom => "bottom",
                        _ when alignment.Vertical.Value == VerticalAlignmentValues.Top => "top",
                        _ => "middle"
                    };
                }
                if (alignment.Indent?.Value != null)
                {
                    result.Styles["padding-inline-start"] = $"{Common.Format(alignment.Indent.Value, configuration)}ch";
                }
                if (alignment.WrapText != null && (alignment.WrapText?.Value ?? true))
                {
                    result.Styles["white-space"] = "preserve wrap";
                }
                if (alignment.TextRotation?.Value != null && alignment.TextRotation.Value != 0)
                {
                    Specification.Html.HtmlAttributes attributes = new()
                    {
                        ["style"] = alignment.TextRotation.Value != 255 ? new Specification.Html.HtmlStyles()
                        {
                            ["display"] = "inline-block",
                            ["rotate"] = alignment.TextRotation.Value > 90 ? $"{Common.Format(alignment.TextRotation.Value - 90, configuration)}deg" : $"-{Common.Format(alignment.TextRotation.Value, configuration)}deg"
                        } : new Specification.Html.HtmlStyles()
                        {
                            ["writing-mode"] = "vertical-rl",
                            ["text-orientation"] = "upright"
                        }
                    };
                    result.Formatters.Add(x => [new Specification.Html.HtmlElement(Specification.Html.HtmlElementType.Paired, "span", attributes, x)]);
                }
                if (alignment.ReadingOrder?.Value != null && alignment.ReadingOrder.Value != 0)
                {
                    result.Styles["direction"] = alignment.ReadingOrder.Value > 1 ? "rtl" : "ltr";
                }
            }

            return result;
        }
    }
}
