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
    /// Initializes a new instance of the <see cref="ConverterProgressCallbackEventArgs"/> class.
    /// </summary>
    public class ConverterProgressCallbackEventArgs((uint Current, uint Total) sheet, (uint Current, uint Total) row) : EventArgs
    {
        /// <summary>
        /// Gets the current progress in percentage.
        /// </summary>
        public double ProgressPercentage { get => Math.Clamp(100.0 * (CurrentSheet + (double)CurrentRow / RowCount - 1) / SheetCount, 0, 100); }

        /// <summary>
        /// Gets the 1-indexed position of the current sheet.
        /// </summary>
        public uint CurrentSheet { get; } = sheet.Current;

        /// <summary>
        /// Gets the total number of sheets.
        /// </summary>
        public uint SheetCount { get; } = sheet.Total;

        /// <summary>
        /// Gets the 1-indexed position of the current row within the current sheet.
        /// </summary>
        public uint CurrentRow { get; } = row.Current;

        /// <summary>
        /// Gets the total number of rows within the current sheet.
        /// </summary>
        public uint RowCount { get; } = row.Total;
    }

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
        /// <param name="callback">The progress callback handler.</param>
        public static void ConvertXlsx(string input, string output, ConverterConfiguration? configuration = null, EventHandler<ConverterProgressCallbackEventArgs>? callback = null)
        {
            configuration ??= new();

            using FileStream stream = new(output, FileMode.Create, FileAccess.Write, FileShare.Read, configuration.BufferSize);
            ConvertXlsx(input, stream, configuration, callback);
        }

        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The path of the local input XLSX document.</param>
        /// <param name="output">The output stream of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback handler.</param>
        public static void ConvertXlsx(string input, Stream output, ConverterConfiguration? configuration = null, EventHandler<ConverterProgressCallbackEventArgs>? callback = null)
        {
            configuration ??= new();

            using FileStream stream = new(input, FileMode.Open, FileAccess.Read, FileShare.Read, configuration.BufferSize);
            ConvertXlsx(stream, output, configuration, callback);
        }

        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The stream of the input XLSX document.</param>
        /// <param name="output">The output stream of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback handler.</param>
        public static void ConvertXlsx(Stream input, Stream output, ConverterConfiguration? configuration = null, EventHandler<ConverterProgressCallbackEventArgs>? callback = null)
        {
            configuration ??= new();

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input, false);
            ConvertXlsx(spreadsheet, output, configuration, callback);
        }

        /// <summary>
        /// Converts a XLSX document to HTML content.
        /// </summary>
        /// <param name="input">The input XLSX document.</param>
        /// <param name="output">The output stream of the HTML content.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <param name="callback">The progress callback handler.</param>
        public static void ConvertXlsx(SpreadsheetDocument input, Stream output, ConverterConfiguration? configuration = null, EventHandler<ConverterProgressCallbackEventArgs>? callback = null)
        {
            configuration ??= new();
            Base.ConverterContext context = new();

            T2 converter<T1, T2>(Base.IConverterBase<T1, T2> converter, T1 value)
            {
                return converter.Convert(value, context, configuration);
            }

            using StreamWriter writer = new(output, configuration.Encoding, configuration.BufferSize, true);

            int indent = 0;
            if (!configuration.UseHtmlFragment)
            {
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Declaration, "html")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "html")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "head")));
                indent++;

                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Unpaired, "meta", new()
                {
                    ["charset"] = "UTF-8"
                })));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Paired, "title", null, [configuration.HtmlTitle])));
            }
            writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Paired, "style", null, [configuration.HtmlPresetStylesheet])));
            if (!configuration.UseHtmlFragment)
            {
                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "head")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "body")));
            }
            indent++;

            WorkbookPart? workbook = input.WorkbookPart;
            context.Theme = workbook?.ThemePart?.Theme;
            context.Stylesheet = converter(configuration.ConverterComposition.XlsxStylesheetReader, workbook?.WorkbookStylesPart?.Stylesheet);
            context.SharedStrings = converter(configuration.ConverterComposition.XlsxSharedStringTableReader, workbook?.SharedStringTablePart?.SharedStringTable);

            IEnumerable<Sheet> sheets = workbook?.Workbook.Sheets?.Elements<Sheet>() ?? [];
            if (!configuration.ConvertHiddenSheets)
            {
                sheets = sheets.Where(x => (x.State ?? SheetStateValues.Visible) == SheetStateValues.Visible);
            }
            if (configuration.ConvertFirstSheetOnly)
            {
                sheets = sheets.Take(1);
            }

            (uint Current, uint Total) progress = (1, (uint)sheets.Count());
            foreach (Sheet sheet in sheets)
            {
                WorksheetPart? worksheet = sheet.Id?.Value != null && (workbook?.TryGetPartById(sheet.Id.Value, out OpenXmlPart? part) ?? false) ? part as WorksheetPart : null;

                context.Worksheet = converter(configuration.ConverterComposition.XlsxWorksheetReader, worksheet?.Worksheet);
                context.Worksheet.Specialties.AddRange(worksheet?.TableDefinitionParts.SelectMany(x => converter(configuration.ConverterComposition.XlsxTableReader, x)) ?? []);
                context.Worksheet.Specialties.AddRange(converter(configuration.ConverterComposition.XlsxDrawingReader, worksheet?.DrawingsPart));

                Dictionary<uint, List<Base.XlsxRangeSpecialty>> references = [];
                foreach (Base.XlsxRangeSpecialty specialty in context.Worksheet.Specialties)
                {
                    for (uint i = specialty.Range.RowStart; i <= specialty.Range.RowEnd; i++)
                    {
                        if (!references.ContainsKey(i))
                        {
                            references[i] = [];
                        }

                        references[i].Add(specialty);
                    }
                }

                double[] lefts = new double[context.Worksheet.Dimension.ColumnCount];
                double tops = 0;
                Dictionary<(uint Column, uint Row), (double Left, double Top)> anchors = [];

                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "table")));
                indent++;

                if (configuration.ConvertSheetTitles)
                {
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Paired, "caption", new()
                    {
                        ["style"] = context.Worksheet.TitleStyles
                    }, [sheet.Name?.Value ?? string.Empty])));
                }

                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "colgroup")));
                indent++;

                double[] widths = new double[context.Worksheet.Dimension.ColumnCount];
                for (uint i = 0; i < widths.Length; i++)
                {
                    lefts[i] = i > 0 ? lefts[i - 1] + widths[i - 1] : 0;
                    widths[i] = Base.Defaults.Common.Get(context.Worksheet.ColumnWidths, i) ?? context.Worksheet.DefaultCellSize.Width;
                }

                double sum = widths.Sum();
                for (uint i = 0; i < widths.Length; i++)
                {
                    lefts[i] = 100.0 * lefts[i] / sum;

                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Unpaired, "col", configuration.ConvertSizes ? new()
                    {
                        ["style"] = new Base.HtmlStyles()
                        {
                            ["width"] = $"{Base.Defaults.Common.Format(100.0 * widths[i] / sum, configuration)}%"
                        }
                    } : null)));
                }

                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "colgroup")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "tbody")));
                indent++;

                bool isOpen = false;
                IEnumerable<Base.XlsxRangeSpecialty> specialties = [];
                (uint Column, uint Row) last = (context.Worksheet.Dimension.ColumnStart - 1, context.Worksheet.Dimension.RowStart - 1);

                void content(uint column, uint row, Base.HtmlAttributeCollection? attributes = null, List<object>? content = null)
                {
                    if (specialties.Any(x => x.Specialty is MergeCell && x.Range.ContainsColumn(column) && !x.Range.StartsAt(column, row)))
                    {
                        return;
                    }

                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Paired, "td", attributes, content)));
                }
                void gap()
                {
                    if (!isOpen)
                    {
                        return;
                    }

                    for (uint i = last.Column + 1; i <= context.Worksheet.Dimension.ColumnEnd; i++)
                    {
                        content(i, last.Row);
                    }

                    indent--;
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "tr")));
                }

                foreach (Row row in context.Worksheet.Data?.Elements<Row>() ?? [])
                {
                    uint index = row.RowIndex?.Value ?? (last.Row + 1);

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        (uint Column, uint Row) current = cell.CellReference?.Value != null ? Base.XlsxRange.ParseReference(cell.CellReference.Value) : (index != last.Row ? context.Worksheet.Dimension.ColumnStart : last.Column + 1, index);
                        if (!context.Worksheet.Dimension.Contains(current.Column, current.Row))
                        {
                            continue;
                        }

                        while (current.Row > last.Row)
                        {
                            gap();

                            double height = (Base.Defaults.Common.Get(row.Height?.Value, current.Row <= last.Row + 1 ? row.CustomHeight?.Value : false) / 72.0 * 96.0) ?? context.Worksheet.DefaultCellSize.Height;
                            writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "tr", configuration.ConvertSizes ? new()
                            {
                                ["style"] = new Base.HtmlStyles()
                                {
                                    ["height"] = $"{Base.Defaults.Common.Format(height, configuration)}px"
                                }
                            } : null)));
                            indent++;

                            isOpen = true;
                            last = (context.Worksheet.Dimension.ColumnStart - 1, last.Row + 1);
                            specialties = Base.Defaults.Common.Get(references, last.Row) ?? [];

                            foreach (Base.XlsxRangeSpecialty specialty in specialties.Where(x => x.Specialty is Base.HtmlElement))
                            {
                                if (specialty.Range.RowStart == last.Row)
                                {
                                    anchors[(specialty.Range.ColumnStart, specialty.Range.RowStart)] = (Base.Defaults.Common.Get(lefts, specialty.Range.ColumnStart - context.Worksheet.Dimension.ColumnStart), tops);
                                }
                                if (specialty.Range.RowEnd == last.Row)
                                {
                                    anchors[(specialty.Range.ColumnEnd, specialty.Range.RowEnd)] = (Base.Defaults.Common.Get(lefts, specialty.Range.ColumnEnd - context.Worksheet.Dimension.ColumnStart), tops);
                                }
                            }
                            tops += height;
                        }
                        if (current.Row != last.Row)
                        {
                            specialties = Base.Defaults.Common.Get(references, current.Row) ?? [];
                        }
                        for (uint i = last.Column + 1; i < current.Column; i++)
                        {
                            content(i, current.Row);
                        }

                        uint? style = cell.StyleIndex?.Value ?? row.StyleIndex?.Value;
                        Base.XlsxContent value = converter(configuration.ConverterComposition.XlsxCellContentReader, new(cell)
                        {
                            NumberFormatId = Base.Defaults.Common.Get(context.Stylesheet.CellFormats, style)?.NumberFormatId ?? 0,
                            Specialties = [.. specialties.Where(x => x.Range.ContainsColumn(current.Column))],
                        });

                        Base.HtmlAttributeCollection attributes = [];
                        List<string> classes = [];

                        if (specialties.FirstOrDefault(x => x.Specialty is MergeCell && x.Range.StartsAt(current.Column, current.Row)) is Base.XlsxRangeSpecialty merge)
                        {
                            attributes["colspan"] = merge.Range.ColumnCount;
                            attributes["rowspan"] = merge.Range.RowCount;
                        }

                        if (configuration.ConvertStyles)
                        {
                            if (style != null && Base.Defaults.Common.Get(context.Stylesheet.CellFormats, style) is Base.XlsxCellFormat format)
                            {
                                value.Styles.Merge(format.Styles, configuration.UseHtmlClasses);

                                if (configuration.UseHtmlClasses)
                                {
                                    classes.Add($"format-{style}");
                                }
                            }

                            foreach (Base.XlsxStyles? styles in specialties.Select(x => x.Specialty is Base.XlsxStyles styles && x.Range.ContainsColumn(current.Column) ? styles : null))
                            {
                                if (styles == null)
                                {
                                    continue;
                                }

                                value.Styles.Merge(styles);
                            }

                            foreach (uint differential in value.DifferentialFormatIds)
                            {
                                value.Styles.Merge((Base.Defaults.Common.Get(context.Stylesheet.DifferentialFormats, differential) ?? new()).Combine(), configuration.UseHtmlClasses);

                                if (configuration.UseHtmlClasses)
                                {
                                    classes.Add($"differential-{differential}");
                                }
                            }

                            attributes["class"] = string.Join(' ', classes);
                            attributes["style"] = value.Styles.Styles;
                            value.Content = value.Styles.ApplyContainers(value.Content);
                        }

                        content(current.Column, current.Row, attributes, value.Content);

                        last = current;
                    }

                    callback?.Invoke(input, new(progress, (last.Row - context.Worksheet.Dimension.RowStart + 1, context.Worksheet.Dimension.RowCount)));
                }
                gap();

                IEnumerable<Base.XlsxRangeSpecialty> elements = context.Worksheet.Specialties.Where(x => x.Specialty is Base.HtmlElement);
                if (elements.Any())
                {
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "tr", new()
                    {
                        ["style"] = new Base.HtmlStyles()
                        {
                            ["visibility"] = "collapse"
                        }
                    })));
                    indent++;

                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedStart, "td")));
                    indent++;

                    foreach (Base.XlsxRangeSpecialty specialty in elements)
                    {
                        if (specialty.Specialty is not Base.HtmlElement element)
                        {
                            continue;
                        }

                        (double left, double top) = Base.Defaults.Common.Get(anchors, (specialty.Range.ColumnStart, specialty.Range.RowStart));
                        (double right, double bottom) = Base.Defaults.Common.Get(anchors, (specialty.Range.ColumnEnd, specialty.Range.RowEnd));
                        Base.HtmlStyles positions = new()
                        {
                            ["--left"] = $"{Base.Defaults.Common.Format(left, configuration)}%",
                            ["--top"] = $"{Base.Defaults.Common.Format(top, configuration)}px",
                            ["--right"] = $"{Base.Defaults.Common.Format(right, configuration)}%",
                            ["--bottom"] = $"{Base.Defaults.Common.Format(bottom, configuration)}px",
                            ["visibility"] = "visible"
                        };

                        element.Indent = indent;
                        if (Base.Defaults.Common.Get(element.Attributes, "style") is Base.HtmlStyles styles)
                        {
                            styles.Merge(positions);
                        }
                        else
                        {
                            element.Attributes["style"] = positions;
                        }

                        writer.Write(converter(configuration.ConverterComposition.HtmlWriter, element));
                    }

                    indent--;
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "td")));

                    indent--;
                    writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "tr")));
                }

                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "tbody")));

                indent--;
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "table")));

                progress = (progress.Current + 1, progress.Total);
            }

            if (configuration.ConvertStyles && configuration.UseHtmlClasses)
            {
                Base.HtmlStylesheetCollection collection = [];
                for (int i = 0; i < context.Stylesheet.CellFormats.Length; i++)
                {
                    collection[$".format-{i}"] = context.Stylesheet.CellFormats[i].Styles.Styles;
                }
                for (int i = 0; i < context.Stylesheet.DifferentialFormats.Length; i++)
                {
                    collection[$".differential-{i}"] = context.Stylesheet.DifferentialFormats[i].Combine().Styles;
                }
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.Paired, "style", null, [collection])));
            }

            indent--;
            if (!configuration.UseHtmlFragment)
            {
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "body")));
                writer.Write(converter(configuration.ConverterComposition.HtmlWriter, new(indent, Base.HtmlElement.ElementType.PairedEnd, "html")));
            }
        }
    }
}

namespace XlsxToHtmlConverter.Base.Defaults
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Common"/> class.
    /// </summary>
    public class Common()
    {
        /// <summary>
        /// Retrieves a specified value.
        /// </summary>
        /// <typeparam name="T">The type of the value.</typeparam>
        /// <param name="value">The specified value.</param>
        /// <param name="flag">Whether the value can be retrieved.</param>
        /// <returns>The retrieved value.</returns>
        public static T? Get<T>(T? value, BooleanValue? flag)
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
        public static T? Get<T>(T?[] values, UInt32Value? index, BooleanValue? flag = null)
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
        public static T2? Get<T1, T2>(Dictionary<T1, T2> values, T1? key, BooleanValue? flag = null) where T1 : notnull
        {
            return key != null && values.TryGetValue(key, out T2? result) ? Get(result, flag) : default;
        }

        /// <summary>
        /// Applies a value to a formatter.
        /// </summary>
        /// <typeparam name="T">The type of the value.</typeparam>
        /// <param name="value">The value.</param>
        /// <param name="formatter">The formatter.</param>
        /// <param name="fallback">The fallback result when the specified value is <see langword="null"/>.</param>
        /// <returns>The formatted result.</returns>
        public static string Use<T>(T? value, Func<T, string> formatter, string? fallback = null) where T : struct
        {
            return value != null ? formatter(value.Value) : (fallback ?? string.Empty);
        }

        /// <summary>
        /// Formats a numeric value.
        /// </summary>
        /// <param name="value">The numeric value.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <returns>The formatted result.</returns>
        public static string Format(double value, ConverterConfiguration configuration)
        {
            return value.ToString(configuration.RoundingDigits < 0 ? "G" : $"F{configuration.RoundingDigits}", CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultHtmlWriter"/> class.
    /// </summary>
    public class DefaultHtmlWriter() : IConverterBase<HtmlElement, string>
    {
        public string Convert(HtmlElement value, ConverterContext context, ConverterConfiguration configuration)
        {
            string padding(int indent)
            {
                return new string(' ', indent * 4);
            }
            string element(HtmlElement element)
            {
                return element.Type switch
                {
                    HtmlElement.ElementType.Declaration => $"<!DOCTYPE {element.Tag}>",
                    HtmlElement.ElementType.Paired => $"<{element.Tag}{attributes(element.Attributes)}>{content(element.Content, element.Indent)}</{element.Tag}>",
                    HtmlElement.ElementType.PairedStart => $"<{element.Tag}{attributes(element.Attributes)}>",
                    HtmlElement.ElementType.PairedEnd => $"</{element.Tag}>",
                    HtmlElement.ElementType.Unpaired => $"<{element.Tag}{attributes(element.Attributes)}>",
                    _ => $"<!-- {content(element.Content, element.Indent)} -->"
                };
            }
            string attributes(HtmlAttributeCollection attributes)
            {
                return string.Concat(attributes.Select(x => x.Value switch
                {
                    null => $" {x.Key}",
                    HtmlStyles styles => styles.Any() ? $" {x.Key}=\"{string.Join(' ', styles.Select(y => $"{y.Key}: {y.Value};"))}\"" : string.Empty,
                    _ => $" {x.Key}=\"{x.Value}\""
                }));
            }
            string content(List<object> content, int indent)
            {
                return string.Concat(content.Select(x =>
                {
                    switch (x)
                    {
                        case HtmlElement html:
                            return element(html);
                        case HtmlStylesheetCollection css:
                            StringBuilder builder = new(configuration.NewlineCharacter);

                            foreach ((string selector, HtmlStyles styles) in css)
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

            return $"{padding(value.Indent)}{element(value)}{configuration.NewlineCharacter}";
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxStylesheetReader"/> class.
    /// </summary>
    public class DefaultXlsxStylesheetReader() : IConverterBase<Stylesheet?, XlsxStylesheetCollection>
    {
        public XlsxStylesheetCollection Convert(Stylesheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxStylesheetCollection result = new();

            XlsxStyles styles<T>(IConverterBase<T, XlsxStyles> converter, T value)
            {
                return converter.Convert(value, context, configuration);
            }
            (string Code, bool IsDate)? builder((StringBuilder Builder, bool IsDate)? builder)
            {
                if (builder == null)
                {
                    return null;
                }

                return (builder.Value.Builder.ToString(), builder.Value.IsDate);
            }

            if (configuration.ConvertStyles)
            {
                Font?[] fonts = [.. value.Fonts?.Elements().Select(x => x as Font) ?? []];
                Fill?[] fills = [.. value.Fills?.Elements().Select(x => x as Fill) ?? []];
                Border?[] borders = [.. value.Borders?.Elements().Select(x => x as Border) ?? []];

                result.CellFormats = [.. value.CellFormats?.Elements().Select(x =>
                {
                    if (x is not CellFormat format)
                    {
                        return new();
                    }

                    XlsxCellFormat cell = new()
                    {
                        NumberFormatId = Common.Get(format.NumberFormatId?.Value, format.ApplyNumberFormat) ?? 0
                    };

                    cell.Styles.Merge(styles(configuration.ConverterComposition.XlsxFontConverter, Common.Get(fonts, format.FontId, format.ApplyFont)));
                    cell.Styles.Merge(styles(configuration.ConverterComposition.XlsxFillConverter, Common.Get(fills, format.FillId, format.ApplyFill)));
                    cell.Styles.Merge(styles(configuration.ConverterComposition.XlsxBorderConverter, Common.Get(borders, format.BorderId, format.ApplyBorder)));
                    cell.Styles.Merge(styles(configuration.ConverterComposition.XlsxAlignmentConverter, Common.Get(format.Alignment, format.ApplyAlignment)));

                    return cell;
                }) ?? []];

                result.DifferentialFormats = [.. value.DifferentialFormats?.Elements().Select(x =>
                {
                    if (x is not DifferentialFormat format)
                    {
                        return new();
                    }

                    return new XlsxDifferentialFormat()
                    {
                        FontStyles = styles(configuration.ConverterComposition.XlsxFontConverter, format.Font),
                        FillStyles = styles(configuration.ConverterComposition.XlsxFillConverter, format.Fill),
                        BorderStyles = styles(configuration.ConverterComposition.XlsxBorderConverter, format.Border),
                        AlignmentStyles = styles(configuration.ConverterComposition.XlsxAlignmentConverter, format.Alignment)
                    };
                }) ?? []];
            }

            foreach (NumberingFormat format in Common.Get(value.NumberingFormats?.Elements<NumberingFormat>(), configuration.ConvertNumberFormats) ?? [])
            {
                if (format.NumberFormatId?.Value == null || WebUtility.HtmlDecode(format.FormatCode?.Value) is not string code || code.All(char.IsWhiteSpace))
                {
                    continue;
                }

                (StringBuilder Builder, bool IsDate)?[] builders = [(new(), false), null, null, null];

                int section = 0;
                foreach ((int index, char character, bool isEscaped) in XlsxNumberFormat.Escape(code, null, ['[', ']']))
                {
                    if (!isEscaped)
                    {
                        switch (char.ToUpperInvariant(character))
                        {
                            case ';' when section < 3:
                                section++;
                                builders[section] = (new(), false);
                                continue;
                            case 'Y' or 'M' or 'D' or 'H' or 'S' when !(builders[section]?.IsDate ?? false):
                                builders[section] = (builders[section]?.Builder ?? new(), true);
                                break;
                        }
                    }

                    builders[section]?.Builder.Append(character);
                }

                builders[1] ??= (new StringBuilder('-').Append(builders[0]?.Builder), builders[0]?.IsDate ?? false);
                builders[2] ??= builders[0];

                result.NumberFormats[format.NumberFormatId.Value] = new(builder(builders[0]), builder(builders[1]), builder(builders[2]), builder(builders[3]));
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxSharedStringTableReader"/> class.
    /// </summary>
    public class DefaultXlsxSharedStringTableReader() : IConverterBase<SharedStringTable?, XlsxString[]>
    {
        public XlsxString[] Convert(SharedStringTable? value, ConverterContext context, ConverterConfiguration configuration)
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
    public class DefaultXlsxWorksheetReader() : IConverterBase<Worksheet?, XlsxWorksheet>
    {
        public XlsxWorksheet Convert(Worksheet? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxWorksheet result = new();

            double width = 64;
            double height = 20;
            Dictionary<uint, double?> widths = [];
            XlsxRange? dimension = null;

            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case SheetData data:
                        result.Data = data;
                        break;
                    case SheetProperties properties:
                        result.TitleStyles["border-bottom"] = $"thick solid {configuration.ConverterComposition.XlsxColorConverter.Convert(properties.TabColor, context, configuration)}";
                        break;
                    case SheetFormatProperties format when configuration.ConvertSizes:
                        if (format.DefaultColumnWidth?.Value != null)
                        {
                            width = format.DefaultColumnWidth.Value * 7.0 + 5;
                        }
                        if (format.DefaultRowHeight?.Value != null)
                        {
                            height = format.DefaultRowHeight.Value / 72.0 * 96.0;
                        }
                        break;
                    case Columns columns when configuration.ConvertSizes:
                        foreach (Column column in columns.Elements<Column>())
                        {
                            if (column.Min?.Value == null)
                            {
                                continue;
                            }

                            for (uint i = column.Min.Value; i <= (column.Max?.Value ?? column.Min.Value); i++)
                            {
                                widths[i] = (column.Collapsed?.Value ?? false) || (column.Hidden?.Value ?? false) ? 0 : Common.Get(column.Width?.Value, column.CustomWidth?.Value) * 7.0 + 5;
                            }
                        }
                        break;
                    case SheetDimension references when references.Reference?.Value != null:
                        dimension = new(references.Reference.Value);
                        break;
                }
            }

            if (dimension == null)
            {
                dimension = new(1, 1, 1, 1);
                foreach (Cell cell in result.Data?.Elements<Row>().SelectMany(x => x.Elements<Cell>()) ?? [])
                {
                    if (cell.CellReference?.Value == null)
                    {
                        continue;
                    }

                    (uint column, uint row) = XlsxRange.ParseReference(cell.CellReference.Value);
                    dimension.ColumnEnd = Math.Max(dimension.ColumnEnd, column);
                    dimension.RowEnd = Math.Max(dimension.RowEnd, row);
                }
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

                            result.Specialties.Add(new()
                            {
                                Specialty = merge,
                                Range = new(merge.Reference.Value, dimension)
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

                            result.Specialties.Add(new()
                            {
                                Specialty = conditional,
                                Range = new(item, dimension)
                            });
                        }
                        break;
                }
            }

            result.DefaultCellSize = (width, height);

            result.ColumnWidths = new double?[dimension.ColumnCount];
            if (configuration.ConvertSizes)
            {
                for (uint i = 0; i < result.ColumnWidths.Length; i++)
                {
                    result.ColumnWidths[i] = Common.Get(widths, dimension.ColumnStart + i) ?? width;
                }
            }

            result.Dimension = dimension;

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxCellContentReader"/> class.
    /// </summary>
    public class DefaultXlsxCellContentReader() : IConverterBase<XlsxCell?, XlsxContent>
    {
        private class NumberInformation
        {
            public List<string> Tokens { get; set; } = [];
            public int Scaling { get; set; } = 0;
            public bool IsGrouped { get; set; } = false;
            public bool IsFractional { get; set; } = false;
            public int[] Lengths { get; set; } = [0, 0, 0, 0];
        }

        private readonly Dictionary<uint, XlsxNumberFormat> formats = new()
        {
            [1] = new(("0", false), ("0", false), ("0", false), null),
            [2] = new(("0.00", false), ("0.00", false), ("0.00", false), null),
            [3] = new(("#,##0", false), ("#,##0", false), ("#,##0", false), null),
            [4] = new(("#,##0.00", false), ("#,##0.00", false), ("#,##0.00", false), null),
            [9] = new(("0%", false), ("0%", false), ("0%", false), null),
            [10] = new(("0.00%", false), ("0.00%", false), ("0.00%", false), null),
            [11] = new(("0.00E+00", false), ("0.00E+00", false), ("0.00E+00", false), null),
            [12] = new(("# ?/?", false), ("# ?/?", false), ("# ?/?", false), null),
            [13] = new(("# ??/??", false), ("# ??/??", false), ("# ??/??", false), null),
            [14] = new(("mm-dd-yy", true), ("mm-dd-yy", true), ("mm-dd-yy", true), null),
            [15] = new(("d-mmm-yy", true), ("d-mmm-yy", true), ("d-mmm-yy", true), null),
            [16] = new(("d-mmm", true), ("d-mmm", true), ("d-mmm", true), null),
            [17] = new(("mmm-yy", true), ("mmm-yy", true), ("mmm-yy", true), null),
            [18] = new(("h:mm AM/PM", true), ("h:mm AM/PM", true), ("h:mm AM/PM", true), null),
            [19] = new(("h:mm:ss AM/PM", true), ("h:mm:ss AM/PM", true), ("h:mm:ss AM/PM", true), null),
            [20] = new(("h:mm", true), ("h:mm", true), ("h:mm", true), null),
            [21] = new(("h:mm:ss", true), ("h:mm:ss", true), ("h:mm:ss", true), null),
            [22] = new(("m/d/yy h:mm", true), ("m/d/yy h:mm", true), ("m/d/yy h:mm", true), null),
            [37] = new(("#,##0 ", false), ("(#,##0)", false), ("#,##0 ", false), null),
            [38] = new(("#,##0 ", false), ("[Red](#,##0)", false), ("#,##0 ", false), null),
            [39] = new(("#,##0.00", false), ("(#,##0.00)", false), ("#,##0.00", false), null),
            [40] = new(("#,##0.00", false), ("[Red](#,##0.00)", false), ("#,##0.00", false), null),
            [45] = new(("mm:ss", true), ("mm:ss", true), ("mm:ss", true), null),
            [46] = new(("[h]:mm:ss", true), ("[h]:mm:ss", true), ("[h]:mm:ss", true), null),
            [47] = new(("mmss.0", true), ("mmss.0", true), ("mmss.0", true), null),
            [48] = new(("##0.0E+0", false), ("##0.0E+0", false), ("##0.0E+0", false), null),
            [49] = new(("@", false), ("@", false), ("@", false), null)
        };
        private readonly Dictionary<string, string> colors = new()
        {
            ["BLACK"] = "000000",
            ["GREEN"] = "008000",
            ["WHITE"] = "FFFFFF",
            ["BLUE"] = "0000FF",
            ["MAGENTA"] = "FF00FF",
            ["YELLOW"] = "FFFF00",
            ["CYAN"] = "00FFFF",
            ["RED"] = "FF0000"
        };
        private readonly Dictionary<string, Func<double, double, bool>> conditions = new()
        {
            ["="] = (x, y) => x == y,
            ["<>"] = (x, y) => x != y,
            ["<"] = (x, y) => x < y,
            ["<="] = (x, y) => x <= y,
            [">"] = (x, y) => x > y,
            [">="] = (x, y) => x >= y
        };

        public XlsxContent Convert(XlsxCell? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxContent result = new();

            string? alignment = null;

            (List<object> Composite, string Raw) text(XlsxString data)
            {
                return (data.Content, data.Raw);
            }
            (List<object> Composite, string Raw) number(object data, string raw)
            {
                if (!configuration.ConvertNumberFormats)
                {
                    return ([raw], raw);
                }

                (int section, (string Code, bool IsDate)? code) = (Common.Get(context.Stylesheet.NumberFormats, value.NumberFormatId) ?? Common.Get(formats, value.NumberFormatId)) is XlsxNumberFormat format ? data switch
                {
                    double number when number > 0 => (0, format.Positive),
                    double number when number < 0 => (1, format.Negative),
                    double number when number == 0 => (2, format.Zero),
                    _ => (3, format.Text)
                } : (3, null);
                object key = ("numFmt", value.NumberFormatId, section);

                string? currency = null;
                CultureInfo culture = configuration.CurrentCulture;
                if (code != null)
                {
                    int start = 0;

                    string? token = null;
                    foreach ((int index, char character, bool isEscaped) in XlsxNumberFormat.Escape(code.Value.Code))
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

                        if (Common.Get(colors, token) is string color)
                        {
                            result.Styles.Styles["color"] = configuration.ConverterComposition.XlsxColorConverter.Convert(new Color()
                            {
                                Rgb = new(color)
                            }, context, configuration);
                        }
                        else if (Common.Get(conditions, string.Concat(token.TakeWhile(x => x is '=' or '<' or '>'))) is Func<double, double, bool> comparator)
                        {
                            if (data is double number && double.TryParse(string.Concat(token.SkipWhile(x => x is '=' or '<' or '>')), NumberStyles.Float, CultureInfo.InvariantCulture, out double operand) && comparator(number, operand))
                            {
                                result.Styles.Styles.Remove("color");
                            }
                        }
                        else if (token.StartsWith('$'))
                        {
                            string[] identifiers = token.TrimStart('$').Split('-');

                            currency = !identifiers[0].All(char.IsWhiteSpace) ? identifiers[0] : null;
                            if (int.TryParse(identifiers[^1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int locale))
                            {
                                try
                                {
                                    culture = CultureInfo.GetCultureInfo(locale);
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            break;
                        }

                        token = null;
                        start = index + 1;
                    }

                    code = (code.Value.Code[start..], code.Value.IsDate);
                }

                if (code == null || code.Value.Code.Trim().ToUpperInvariant() == "GENERAL")
                {
                    switch (data)
                    {
                        case DateTime date:
                            alignment = "right";

                            return ([time(date, date.ToString("d", culture))], raw);
                        case double number:
                            alignment = "right";

                            string general = number.ToString("G", CultureInfo.InvariantCulture).Replace("+", string.Empty);
                            if (general.Length <= (general.StartsWith('-') ? 12 : 11))
                            {
                                return ([general], raw);
                            }

                            string scientific = number.ToString("0.#######E0", CultureInfo.InvariantCulture);
                            return ([number.ToString($"0.{new string('#', Math.Max(0, (scientific.StartsWith('-') ? 10 : 9) - (scientific.Length - scientific.IndexOf('E'))))}E0", CultureInfo.InvariantCulture)], raw);
                        default:
                            return ([raw], raw);
                    }
                }

                StringBuilder builder = new();
                if (code.Value.IsDate)
                {
                    if (data is double number && number >= -657435.0 && number <= 2958465.99999999)
                    {
                        data = DateTime.FromOADate(number);
                    }
                    if (data is not DateTime date)
                    {
                        return ([raw], raw);
                    }

                    if (Common.Get(context.Cache, key) is not List<string> information)
                    {
                        information = tokens(code.Value.Code, true, (x, y) => x switch
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

                        context.Cache[key] = information;
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
                        information.Tokens = tokens(code.Value.Code, false, (x, y) =>
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
                                case '+' or '-' when y.Length == 1 && (y[0] is 'E' or 'e'):
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
                                    return y.Length == 1 && (y[0] is '_' or '*') ? false : null;
                            }
                        }, ['_', '*']);

                        context.Cache[key] = information;
                    }

                    number *= Math.Pow(10, information.Scaling);

                    (long Numerator, int Denominator)? fraction = null;
                    if (information.IsFractional)
                    {
                        long whole = (long)number;
                        double remainder = number - whole;
                        fraction = remainder switch
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
                            fraction = (fraction.Value.Denominator * whole + fraction.Value.Numerator, fraction.Value.Denominator);
                        }
                        else
                        {
                            number = whole;
                        }
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
                        components[2] = exponent.ToString("D", CultureInfo.InvariantCulture).PadLeft(information.Lengths[2], ' ');
                    }
                    else
                    {
                        number = Math.Round(number, information.Lengths[1]);
                    }

                    long integer = (long)number;
                    components[0] = integer.ToString("D", CultureInfo.InvariantCulture).PadLeft(information.Lengths[0], ' ');
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

                                    string numerator = (fraction?.Numerator ?? 0).ToString("D", CultureInfo.InvariantCulture).PadLeft(left.Length, ' ');
                                    string denominator = (fraction?.Denominator ?? 0).ToString("D", CultureInfo.InvariantCulture).PadRight(right.Length, ' ');

                                    digit(left.PadLeft(numerator.Length, '0'), numerator, 0);
                                    builder.Append('/');
                                    digit(right.PadRight(denominator.Length, '0'), denominator, 0);

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
                            case 'E' or 'e' when stage < 2:
                                stage = 2;
                                index = 0;
                                builder.Append(token.First());

                                if (sign is '-' || token.Length > 1)
                                {
                                    builder.Append(sign);
                                }

                                break;
                            case '%':
                                builder.Append(culture.NumberFormat.PercentSymbol);
                                break;
                            case '_':
                                builder.Append(' ');
                                break;
                            case ',' or '*':
                                break;
                            default:
                                literal(builder, token);
                                break;
                        }
                    }
                }

                alignment = "right";

                return data switch
                {
                    DateTime date => ([time(date, builder.ToString())], raw),
                    _ => ([builder.ToString()], raw)
                };
            }
            object time(DateTime date, string content)
            {
                return new HtmlElement("time", new HtmlAttributeCollection()
                {
                    ["datetime"] = date.ToString("O", CultureInfo.InvariantCulture)
                }, [content]);
            }
            List<string> tokens(string code, bool isUppercase, Func<char, StringBuilder, bool?> tokenizer, char[]? escapers = null)
            {
                StringBuilder builder = new();
                List<string> tokens = [];

                bool isSpecial = false;
                foreach ((int index, char character, bool isEscaped) in XlsxNumberFormat.Escape(code, escapers))
                {
                    if (isEscaped)
                    {
                        builder.Append(character);
                        continue;
                    }

                    char input = isUppercase ? char.ToUpperInvariant(character) : character;
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
                foreach ((int index, char character, bool isEscaped) in XlsxNumberFormat.Escape(content))
                {
                    if (!isEscaped && character is '\\' or '\"')
                    {
                        continue;
                    }

                    builder.Append(character);
                }
            }

            string content = value.Cell.CellValue?.Text ?? string.Empty;
            (List<object> composite, string raw) = value.Cell.DataType?.Value switch
            {
                _ when value.Cell.DataType?.Value == CellValues.Error => ([content], content),
                _ when value.Cell.DataType?.Value == CellValues.String => ([content], content),
                _ when value.Cell.DataType?.Value == CellValues.InlineString => text(configuration.ConverterComposition.XlsxStringConverter.Convert(value.Cell, context, configuration)),
                _ when value.Cell.DataType?.Value == CellValues.SharedString => uint.TryParse(content, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint index) && Common.Get(context.SharedStrings, index) is XlsxString shared ? text(shared) : ([], string.Empty),
                _ when value.Cell.DataType?.Value == CellValues.Boolean => ([content.Trim() switch {
                    "1" => "TRUE",
                    "0" => "FALSE",
                    _ => string.Empty
                }], content),
                _ when value.Cell.DataType?.Value == CellValues.Date => DateTime.TryParse(content, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date) ? number(date, content) : ([content], content),
                _ => number(double.TryParse(content, NumberStyles.Float, CultureInfo.InvariantCulture, out double data) ? data : content, content)
            };

            result.Content = composite;

            alignment ??= value.Cell.DataType?.Value switch
            {
                _ when value.Cell.DataType?.Value == CellValues.Error => "center",
                _ when value.Cell.DataType?.Value == CellValues.Boolean => "center",
                _ => null
            };
            if (alignment != null)
            {
                result.Styles.Styles["text-align"] = alignment;
            }

            foreach (XlsxRangeSpecialty specialty in value.Specialties)
            {
                switch (specialty.Specialty)
                {
                    case ConditionalFormatting conditional when configuration.ConvertStyles:
                        foreach (ConditionalFormattingRule rule in conditional.Elements<ConditionalFormattingRule>().OrderByDescending(x => x.Priority?.Value ?? int.MaxValue))
                        {
                            if (rule.Type?.Value == null || rule.FormatId?.Value == null)
                            {
                                continue;
                            }

                            bool cell(ConditionalFormattingOperatorValues operation)
                            {
                                double? number = double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double data) ? data : null;

                                Formula[] formulas = [.. rule.Elements<Formula>()];
                                double?[] targets = [.. formulas.Select(x => double.TryParse(x.Text.Trim('\"'), NumberStyles.Float, CultureInfo.InvariantCulture, out double target) ? target : number)];

                                return operation switch
                                {
                                    _ when operation == ConditionalFormattingOperatorValues.Equal && formulas.Length > 0 => raw.Equals(formulas[0].Text.Trim('\"'), StringComparison.OrdinalIgnoreCase) || (number != null && targets[0] != null && number == targets[0]),
                                    _ when operation == ConditionalFormattingOperatorValues.NotEqual && formulas.Length > 0 => !raw.Equals(formulas[0].Text.Trim('\"'), StringComparison.OrdinalIgnoreCase) || (number != null && targets[0] != null && number != targets[0]),
                                    _ when operation == ConditionalFormattingOperatorValues.LessThan && formulas.Length > 0 && number != null && targets[0] != null => number < targets[0],
                                    _ when operation == ConditionalFormattingOperatorValues.LessThanOrEqual && formulas.Length > 0 && number != null && targets[0] != null => number <= targets[0],
                                    _ when operation == ConditionalFormattingOperatorValues.GreaterThan && formulas.Length > 0 && number != null && targets[0] != null => number > targets[0],
                                    _ when operation == ConditionalFormattingOperatorValues.GreaterThanOrEqual && formulas.Length > 0 && number != null && targets[0] != null => number >= targets[0],
                                    _ when operation == ConditionalFormattingOperatorValues.Between && formulas.Length > 1 && number != null && targets[0] != null && targets[1] != null => number >= Math.Min(targets[0] ?? 0, targets[1] ?? 0) && number <= Math.Max(targets[0] ?? 0, targets[1] ?? 0),
                                    _ when operation == ConditionalFormattingOperatorValues.NotBetween && formulas.Length > 1 && number != null && targets[0] != null && targets[1] != null => number < Math.Min(targets[0] ?? 0, targets[1] ?? 0) || number > Math.Max(targets[0] ?? 0, targets[1] ?? 0),
                                    _ when operation == ConditionalFormattingOperatorValues.ContainsText && formulas.Length > 0 => raw.Contains(formulas[0].Text.Trim('\"'), StringComparison.OrdinalIgnoreCase),
                                    _ when operation == ConditionalFormattingOperatorValues.NotContains && formulas.Length > 0 => !raw.Contains(formulas[0].Text.Trim('\"'), StringComparison.OrdinalIgnoreCase),
                                    _ when operation == ConditionalFormattingOperatorValues.BeginsWith && formulas.Length > 0 => raw.StartsWith(formulas[0].Text.Trim('\"'), StringComparison.OrdinalIgnoreCase),
                                    _ when operation == ConditionalFormattingOperatorValues.EndsWith && formulas.Length > 0 => raw.EndsWith(formulas[0].Text.Trim('\"'), StringComparison.OrdinalIgnoreCase),
                                    _ => false
                                };
                            }

                            //TODO: conditional formatting

                            bool isSatisfied = rule.Type.Value switch
                            {
                                _ when rule.Type.Value == ConditionalFormatValues.CellIs && rule.Operator?.Value != null => cell(rule.Operator.Value),
                                _ when rule.Type.Value == ConditionalFormatValues.ContainsText && rule.Text?.Value != null => raw.Contains(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.NotContainsText && rule.Text?.Value != null => !raw.Contains(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.BeginsWith && rule.Text?.Value != null => raw.StartsWith(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.EndsWith && rule.Text?.Value != null => raw.EndsWith(rule.Text.Value, StringComparison.OrdinalIgnoreCase),
                                _ when rule.Type.Value == ConditionalFormatValues.ContainsBlanks => raw.Any(char.IsWhiteSpace),
                                _ when rule.Type.Value == ConditionalFormatValues.NotContainsBlanks => !raw.Any(char.IsWhiteSpace),
                                _ when rule.Type.Value == ConditionalFormatValues.ContainsErrors => value.Cell.DataType?.Value == CellValues.Error,
                                _ when rule.Type.Value == ConditionalFormatValues.NotContainsErrors => value.Cell.DataType?.Value != CellValues.Error,
                                _ => false,
                            };

                            if (isSatisfied)
                            {
                                result.DifferentialFormatIds.Add(rule.FormatId.Value);

                                if (rule.StopIfTrue?.Value ?? false)
                                {
                                    break;
                                }
                            }
                        }

                        break;
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxTableReader"/> class.
    /// </summary>
    public class DefaultXlsxTableReader() : IConverterBase<TableDefinitionPart?, XlsxRangeSpecialty[]>
    {
        public XlsxRangeSpecialty[] Convert(TableDefinitionPart? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value?.Table == null || value.Table.Reference?.Value == null || !configuration.ConvertStyles)
            {
                return [];
            }

            List<XlsxRangeSpecialty> result = [];

            XlsxRange range = new(value.Table.Reference.Value, context.Worksheet.Dimension);
            uint start = value.Table.HeaderRowCount?.Value ?? 1;
            uint end = value.Table.TotalsRowCount?.Value ?? 0;
            uint middle = range.RowCount - start - end;

            XlsxStyles? styles(uint? body, uint? boundary, uint count)
            {
                if (count <= 0)
                {
                    return null;
                }

                XlsxStyles? styles = null;

                //TODO: table formats

                if (Common.Get(context.Stylesheet.DifferentialFormats, body) is XlsxDifferentialFormat cell)
                {
                    styles ??= new();
                    styles.Merge(cell.FontStyles);
                    styles.Merge(cell.FillStyles);
                    styles.Merge(cell.AlignmentStyles);
                }
                if (Common.Get(context.Stylesheet.DifferentialFormats, boundary) is XlsxDifferentialFormat border)
                {
                    styles ??= new();
                    styles.Merge(border.BorderStyles);
                }

                return styles;
            }

            XlsxStyles? header = styles(value.Table.HeaderRowFormatId?.Value, value.Table.HeaderRowBorderFormatId?.Value, start);
            if (header != null)
            {
                result.Add(new()
                {
                    Specialty = header,
                    Range = new(range.ColumnStart, range.RowStart, range.ColumnEnd, range.RowStart + start - 1)
                });
            }

            XlsxStyles? body = styles(value.Table.DataFormatId?.Value, value.Table.BorderFormatId?.Value, middle);
            if (body != null)
            {
                result.Add(new()
                {
                    Specialty = body,
                    Range = new(range.ColumnStart, range.RowStart + start, range.ColumnEnd, range.RowEnd - end)
                });
            }

            XlsxStyles? totals = styles(value.Table.TotalsRowFormatId?.Value, value.Table.TotalsRowBorderFormatId?.Value, end);
            if (totals != null)
            {
                result.Add(new()
                {
                    Specialty = totals,
                    Range = new(range.ColumnStart, range.RowEnd - end + 1, range.ColumnEnd, range.RowEnd)
                });
            }

            return [.. result];
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxDrawingReader"/> class.
    /// </summary>
    public class DefaultXlsxDrawingReader : IConverterBase<DrawingsPart?, XlsxRangeSpecialty[]>
    {
        public XlsxRangeSpecialty[] Convert(DrawingsPart? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return [];
            }

            return [..value.WorksheetDrawing.Elements().Select(x =>
            {
                HtmlStyles styles = new()
                {
                    ["position"] = "absolute"
                };
                HtmlElement element = new("div", new HtmlAttributeCollection()
                {
                    ["style"] = styles
                });

                uint left = context.Worksheet.Dimension.ColumnStart;
                uint top = context.Worksheet.Dimension.RowStart;
                uint right = context.Worksheet.Dimension.ColumnStart;
                uint bottom = context.Worksheet.Dimension.RowStart;

                static double offset(string? text)
                {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint offset) ? offset / 914400.0 * 96.0 : 0;
                }

                switch (x)
                {
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor absolute:
                        styles["left"] = $"calc(var(--left) + {Common.Format((absolute.Position?.X?.Value / 914400.0 * 96.0) ?? 0, configuration)}px)";
                        styles["top"] = $"calc(var(--top) + {Common.Format((absolute.Position?.Y?.Value / 914400.0 * 96.0) ?? 0, configuration)}px)";
                        styles["width"] = $"{Common.Format((absolute.Extent?.Cx?.Value / 914400.0 * 96.0) ?? 0, configuration)}px";
                        styles["height"] = $"{Common.Format((absolute.Extent?.Cy?.Value / 914400.0 * 96.0) ?? 0, configuration)}px";
                        break;
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor single:
                        if (uint.TryParse(single.FromMarker?.ColumnId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out left))
                        {
                            right = left;
                        }
                        if (uint.TryParse(single.FromMarker?.RowId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out top))
                        {
                            bottom = top;
                        }

                        styles["left"] = $"calc(var(--left) + {Common.Format(offset(single.FromMarker?.ColumnOffset?.Text), configuration)}px)";
                        styles["top"] = $"calc(var(--top) + {Common.Format(offset(single.FromMarker?.RowOffset?.Text), configuration)}px)";
                        styles["width"] = $"{Common.Format((single.Extent?.Cx?.Value / 914400.0 * 96.0) ?? 0, configuration)}px";
                        styles["height"] = $"{Common.Format((single.Extent?.Cy?.Value / 914400.0 * 96.0) ?? 0, configuration)}px";

                        break;
                    case DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor dual:
                        uint.TryParse(dual.FromMarker?.ColumnId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out left);
                        uint.TryParse(dual.FromMarker?.RowId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out top);
                        uint.TryParse(dual.ToMarker?.ColumnId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out right);
                        uint.TryParse(dual.ToMarker?.RowId?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out bottom);

                        double[] offsets = [offset(dual.FromMarker?.ColumnOffset?.Text), offset(dual.FromMarker?.RowOffset?.Text), offset(dual.ToMarker?.ColumnOffset?.Text), offset(dual.ToMarker?.RowOffset?.Text)];
                        styles["left"] = $"calc(var(--left) + {Common.Format(offsets[0], configuration)}px)";
                        styles["top"] = $"calc(var(--top) + {Common.Format(offsets[1], configuration)}px)";
                        styles["width"] = $"calc(var(--right) + {Common.Format(offsets[2], configuration)}px - var(--left) - {Common.Format(offsets[0], configuration)}px)";
                        styles["height"] = $"calc(var(--bottom) + {Common.Format(offsets[3], configuration)}px - var(--top) - {Common.Format(offsets[1], configuration)}px)";

                        break;
                }

                foreach (OpenXmlElement child in x.Elements())
                {
                    switch (child)
                    {
                        case DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture when configuration.ConvertPictures:
                            HtmlAttributeCollection image = new()
                            {
                                ["loading"] = "lazy",
                                ["decoding"] = "async"
                            };

                            if (picture.BlipFill?.Blip?.Embed?.Value != null && value.TryGetPartById(picture.BlipFill.Blip.Embed.Value, out OpenXmlPart? part) && part is ImagePart source)
                            {
                                using MemoryStream memory = new();
                                using Stream stream = source.GetStream();
                                stream.CopyTo(memory);

                                image["src"] = $"data:{source.ContentType};base64,{System.Convert.ToBase64String(memory.ToArray())}";
                            }
                            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value != null)
                            {
                                image["alt"] = WebUtility.HtmlEncode(picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description.Value);
                            }

                            element.Content.Add(new HtmlElement(HtmlElement.ElementType.Unpaired, "img", image));

                            //TODO: pictures

                            break;
                        case DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape shape when configuration.ConvertShapes:

                            //TODO: shapes

                            break;
                    }
                }

                return new XlsxRangeSpecialty()
                {
                    Specialty = element,
                    Range = new(left, top, right, bottom)
                };
            })];
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxColorConverter"/> class.
    /// </summary>
    public class DefaultXlsxColorConverter() : IConverterBase<OpenXmlElement?, string>
    {
        private readonly (byte Red, byte Green, byte Blue)[] indexes = [
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
        private readonly Dictionary<DocumentFormat.OpenXml.Drawing.SystemColorValues, (byte Red, byte Green, byte Blue)> systems = new()
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
        private readonly Dictionary<DocumentFormat.OpenXml.Drawing.PresetColorValues, (byte Red, byte Green, byte Blue)> presets = new()
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
                string formatted = hex.TrimStart('#').PadLeft(8, 'F');
                alpha = byte.TryParse(formatted[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte first) ? first : byte.MaxValue;
                red = byte.TryParse(formatted[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte second) ? second : byte.MinValue;
                green = byte.TryParse(formatted[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte third) ? third : byte.MinValue;
                blue = byte.TryParse(formatted[6..8], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte fourth) ? fourth : byte.MinValue;
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

                red = Math.Clamp(rgb[0] * 255.0, 0, 255);
                green = Math.Clamp(rgb[1] * 255.0, 0, 255);
                blue = Math.Clamp(rgb[2] * 255.0, 0, 255);
            }
            bool element(OpenXmlElement color, IEnumerable<OpenXmlElement> children)
            {
                switch (color)
                {
                    case DocumentFormat.OpenXml.Drawing.RgbColorModelHex model when model.Val?.Value != null:
                        hex(model.Val.Value);
                        break;
                    case DocumentFormat.OpenXml.Drawing.RgbColorModelPercentage percentage:
                        red = Math.Clamp((percentage.RedPortion?.Value ?? 0) / 100000.0 * 255.0, 0, 255);
                        green = Math.Clamp((percentage.GreenPortion?.Value ?? 0) / 100000.0 * 255.0, 0, 255);
                        blue = Math.Clamp((percentage.BluePortion?.Value ?? 0) / 100000.0 * 255.0, 0, 255);
                        break;
                    case DocumentFormat.OpenXml.Drawing.HslColor hsl:
                        modifier(x => (hsl.HueValue?.Value ?? 0) / 60000.0, x => (hsl.SatValue?.Value ?? 0) / 60000.0, x => (hsl.LumValue?.Value ?? 0) / 60000.0);
                        break;
                    case DocumentFormat.OpenXml.Drawing.SystemColor system when system.Val?.Value != null && systems.ContainsKey(system.Val.Value):
                        red = systems[system.Val.Value].Red;
                        green = systems[system.Val.Value].Green;
                        blue = systems[system.Val.Value].Blue;
                        break;
                    case DocumentFormat.OpenXml.Drawing.SystemColor system when system.LastColor?.Value != null:
                        hex(system.LastColor.Value);
                        break;
                    case DocumentFormat.OpenXml.Drawing.PresetColor preset when preset.Val?.Value != null && presets.ContainsKey(preset.Val.Value):
                        red = presets[preset.Val.Value].Red;
                        green = presets[preset.Val.Value].Green;
                        blue = presets[preset.Val.Value].Blue;
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
                        case DocumentFormat.OpenXml.Drawing.Shade shade when (shade.Val?.Value / 100000.0) is double number:
                            red = Math.Clamp(red * number, 0, 255);
                            green = Math.Clamp(green * number, 0, 255);
                            blue = Math.Clamp(blue * number, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Tint tint when (tint.Val?.Value / 100000.0) is double number:
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
                            red = Math.Clamp((red / 255.0 > 0.04045 ? Math.Pow((red / 255.0 + 0.055) / 1.055, 2.4) : red / 255.0 / 12.92) * 255.0, 0, 255);
                            green = Math.Clamp((green / 255.0 > 0.04045 ? Math.Pow((green / 255.0 + 0.055) / 1.055, 2.4) : green / 255.0 / 12.92) * 255.0, 0, 255);
                            blue = Math.Clamp((blue / 255.0 > 0.04045 ? Math.Pow((blue / 255.0 + 0.055) / 1.055, 2.4) : blue / 255.0 / 12.92) * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.InverseGamma:
                            red = Math.Clamp((red / 255.0 > 0.0031308 ? 1.055 * Math.Pow(red / 255.0, 1 / 2.4) - 0.055 : red / 255.0 * 12.92) * 255.0, 0, 255);
                            green = Math.Clamp((green / 255.0 > 0.0031308 ? 1.055 * Math.Pow(green / 255.0, 1 / 2.4) - 0.055 : green / 255.0 * 12.92) * 255.0, 0, 255);
                            blue = Math.Clamp((blue / 255.0 > 0.0031308 ? 1.055 * Math.Pow(blue / 255.0, 1 / 2.4) - 0.055 : blue / 255.0 * 12.92) * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Red channel when channel.Val?.Value != null:
                            red = Math.Clamp(channel.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.RedModulation modulation when modulation.Val?.Value != null:
                            red = Math.Clamp(red * (modulation.Val.Value / 100000.0), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.RedOffset offset when offset.Val?.Value != null:
                            red = Math.Clamp(red + offset.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Green channel when channel.Val?.Value != null:
                            green = Math.Clamp(channel.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.GreenModulation modulation when modulation.Val?.Value != null:
                            green = Math.Clamp(green * (modulation.Val.Value / 100000.0), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.GreenOffset offset when offset.Val?.Value != null:
                            green = Math.Clamp(green + offset.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Blue channel when channel.Val?.Value != null:
                            blue = Math.Clamp(channel.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.BlueModulation modulation when modulation.Val?.Value != null:
                            blue = Math.Clamp(blue * (modulation.Val.Value / 100000.0), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.BlueOffset offset when offset.Val?.Value != null:
                            blue = Math.Clamp(blue + offset.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Alpha channel when channel.Val?.Value != null:
                            alpha = Math.Clamp(channel.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.AlphaModulation modulation when modulation.Val?.Value != null:
                            alpha = Math.Clamp(alpha * (modulation.Val.Value / 100000.0), 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.AlphaOffset offset when offset.Val?.Value != null:
                            alpha = Math.Clamp(alpha + offset.Val.Value / 100000.0 * 255.0, 0, 255);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Hue channel when channel.Val?.Value != null:
                            modifier(x => channel.Val.Value / 60000.0, x => x, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.HueModulation modulation when modulation.Val?.Value != null:
                            modifier(x => x * (modulation.Val.Value / 100000.0), x => x, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.HueOffset offset when offset.Val?.Value != null:
                            modifier(x => x + offset.Val.Value / 60000.0, x => x, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Saturation channel when channel.Val?.Value != null:
                            modifier(x => x, x => channel.Val.Value / 100000.0, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.SaturationModulation modulation when modulation.Val?.Value != null:
                            modifier(x => x, x => x * (modulation.Val.Value / 100000.0), x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.SaturationOffset offset when offset.Val?.Value != null:
                            modifier(x => x, x => x + offset.Val.Value / 100000.0, x => x);
                            break;
                        case DocumentFormat.OpenXml.Drawing.Luminance channel when channel.Val?.Value != null:
                            modifier(x => x, x => x, x => channel.Val.Value / 100000.0);
                            break;
                        case DocumentFormat.OpenXml.Drawing.LuminanceModulation modulation when modulation.Val?.Value != null:
                            modifier(x => x, x => x, x => x * (modulation.Val.Value / 100000.0));
                            break;
                        case DocumentFormat.OpenXml.Drawing.LuminanceOffset offset when offset.Val?.Value != null:
                            modifier(x => x, x => x, x => x + offset.Val.Value / 100000.0);
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
                else if (color.Indexed?.Value != null && color.Indexed.Value < indexes.Length)
                {
                    red = indexes[color.Indexed.Value].Red;
                    green = indexes[color.Indexed.Value].Green;
                    blue = indexes[color.Indexed.Value].Blue;
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
                (false, false) => $"rgb({result[0]} {result[1]} {result[2]})",
                (false, true) => $"rgb({result[0]} {result[1]} {result[2]} / {Common.Format(result[3] / 255.0, configuration)})",
                (true, false) => $"#{result[0]:X2}{result[1]:X2}{result[2]:X2}",
                _ => $"#{result[0]:X2}{result[1]:X2}{result[2]:X2}{result[3]:X2}",
            };
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxStringConverter"/> class.
    /// </summary>
    public class DefaultXlsxStringConverter : IConverterBase<OpenXmlElement?, XlsxString>
    {
        public XlsxString Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxString result = new();

            StringBuilder builder = new();

            foreach (OpenXmlElement child in value.Elements())
            {
                switch (child)
                {
                    case Text text:
                        builder.Append(text.Text);
                        result.Content.Add(text.Text);
                        break;
                    case Run run when run.Text?.Text != null:
                        builder.Append(run.Text.Text);
                        if (configuration.ConvertStyles)
                        {
                            XlsxStyles format = configuration.ConverterComposition.XlsxFontConverter.Convert(run.RunProperties, context, configuration);
                            result.Content.Add(new HtmlElement("span", format, [run.Text.Text]));
                        }
                        else
                        {
                            result.Content.Add(run.Text.Text);
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
    public class DefaultXlsxFontConverter : IConverterBase<OpenXmlElement?, XlsxStyles>
    {
        public XlsxStyles Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxStyles result = new();

            string color(OpenXmlElement? color)
            {
                return configuration.ConverterComposition.XlsxColorConverter.Convert(color, context, configuration);
            }

            List<string> decorations = [];

            if (value is DocumentFormat.OpenXml.Drawing.TextCharacterPropertiesType properties)
            {
                if (properties.FontSize?.Value != null)
                {
                    result.Styles["font-size"] = $"{Common.Format(properties.FontSize.Value / 7200.0 * 96.0, configuration)}px";
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
                        result.Containers.Add(new()
                        {
                            ["display"] = "inline-block",
                            ["text-decoration"] = "line-through double"
                        });
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
                        result.Containers.Add(new()
                        {
                            ["display"] = "inline-block",
                            ["text-decoration"] = $"underline {properties.Underline.Value switch
                            {
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Double => "double",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dash => "dashed",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLong => "dashed",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDash => "dashed",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashHeavy => "dashed 4px",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DashLongHeavy => "dashed 4px",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDashHeavy => "dashed 4px",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Dotted => "dotted",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDash => "dotted",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.HeavyDotted => "dotted 4px",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.DotDotDashHeavy => "dotted 4px",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Wavy => "wavy",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyDouble => "wavy 4px",
                                _ when properties.Underline.Value == DocumentFormat.OpenXml.Drawing.TextUnderlineValues.WavyHeavy => "wavy 4px",
                                _ => "4px"
                            }}"
                        });
                    }
                    result.Styles["text-decoration"] = "none";
                }
                if (properties.Spacing?.Value != null)
                {
                    result.Styles["letter-spacing"] = $"{Common.Format(properties.Spacing.Value / 7200.0 * 96.0, configuration)}px";
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
                        result.Styles["font-size"] = $"{Common.Format(size.Val.Value / 72.0 * 96.0, configuration)}px";
                        break;
                    case RunFont name when name.Val?.Value != null:
                        result.Styles["font-family"] = $"\'{name.Val.Value}\'";
                        break;
                    case FontName name when name.Val?.Value != null:
                        result.Styles["font-family"] = $"\'{name.Val.Value}\'";
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
                            result.Containers.Add(new()
                            {
                                ["display"] = "inline-block",
                                ["text-decoration"] = "underline double"
                            });
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
                        result.Styles["font-family"] = $"\'{name.Typeface.Value}\'";
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
    public class DefaultXlsxFillConverter : IConverterBase<OpenXmlElement?, XlsxStyles>
    {
        public XlsxStyles Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxStyles result = new();

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
                    result.Styles.Merge(fill.PatternFill.PatternType?.Value switch
                    {
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkGray => new()
                        {
                            ["background-image"] = $"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 2.6px 2.6px, {foreground} 0.5px, transparent 0)",
                            ["background-size"] = "3.2px 3.2px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.MediumGray => new()
                        {
                            ["background-image"] = $"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 2.8px 2.8px, {foreground} 0.5px, transparent 0)",
                            ["background-size"] = "3.6px 3.6px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightGray => new()
                        {
                            ["background-image"] = $"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 3px 3px, {foreground} 0.5px, transparent 0)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.Gray125 => new()
                        {
                            ["background-image"] = $"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 4px 4px, {foreground} 0.5px, transparent 0)",
                            ["background-size"] = "6px 6px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.Gray0625 => new()
                        {
                            ["background-image"] = $"radial-gradient(circle at 1px 1px, {foreground} 0.5px, transparent 0), radial-gradient(circle at 5.5px 5.5px, {foreground} 0.5px, transparent 0)",
                            ["background-size"] = "9px 9px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkHorizontal => new()
                        {
                            ["background-image"] = $"linear-gradient(0deg, {foreground} 1.5px, transparent 0)",
                            ["background-size"] = "100% 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightHorizontal => new()
                        {
                            ["background-image"] = $"linear-gradient(0deg, {foreground} 1px, transparent 0)",
                            ["background-size"] = "100% 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkVertical => new()
                        {
                            ["background-image"] = $"linear-gradient(90deg, {foreground} 1.5px, transparent 0)",
                            ["background-size"] = "4px 100%"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightVertical => new()
                        {
                            ["background-image"] = $"linear-gradient(90deg, {foreground} 1px, transparent 0)",
                            ["background-size"] = "4px 100%"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkDown => new()
                        {
                            ["background-image"] = $"linear-gradient(45deg, {foreground} 25%, transparent 25% 50%, {foreground} 50% 75%, transparent 75%)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightDown => new()
                        {
                            ["background-image"] = $"linear-gradient(45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkUp => new()
                        {
                            ["background-image"] = $"linear-gradient(-45deg, {foreground} 25%, transparent 25% 50%, {foreground} 50% 75%, transparent 75%)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightUp => new()
                        {
                            ["background-image"] = $"linear-gradient(-45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkGrid => new()
                        {
                            ["background-image"] = $"linear-gradient(45deg, {foreground} 25%, transparent 25% 75%, {foreground} 75%), linear-gradient(45deg, {foreground} 25%, transparent 25% 75%, {foreground} 75%)",
                            ["background-position"] = "0 0, 2.5px 2.5px",
                            ["background-size"] = "5px 5px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightGrid => new()
                        {
                            ["background-image"] = $"linear-gradient(90deg, {foreground} 1px, transparent 0), linear-gradient(0deg, {foreground} 1px, transparent 0)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.DarkTrellis => new()
                        {
                            ["background-image"] = $"linear-gradient(45deg, {foreground} 15%, transparent 15% 50%, {foreground} 50% 65%, transparent 65%), linear-gradient(-45deg, {foreground} 15%, transparent 15% 50%, {foreground} 50% 65%, transparent 65%)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.PatternType?.Value == PatternValues.LightTrellis => new()
                        {
                            ["background-image"] = $"linear-gradient(45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%), linear-gradient(-45deg, {foreground} 10%, transparent 10% 50%, {foreground} 50% 60%, transparent 60%)",
                            ["background-size"] = "4px 4px"
                        },
                        _ when fill.PatternFill.ForegroundColor != null => new()
                        {
                            ["background"] = foreground
                        },
                        _ => []
                    });
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
                        result.Styles["background"] = $"radial-gradient(circle at {Common.Format(100.0 * (left + right) / 2, configuration)}% {Common.Format(100.0 * (top + bottom) / 2, configuration)}%{string.Concat(stops.Select(x => $", {color(x.Color)}{Common.Use(x.Position?.Value, y => $" {Common.Format(100.0 * (radius + y * (1 - radius)), configuration)}%")}"))})";
                    }
                    else
                    {
                        double degree = (((fill.GradientFill.Degree?.Value ?? 0) + 90) % 360 + 360) % 360;

                        IEnumerable<GradientStop> stops = fill.GradientFill.Elements<GradientStop>();
                        result.Styles["background"] = $"linear-gradient({Common.Format(degree, configuration)}deg{string.Concat(stops.Select(x => $", {color(x.Color)}{Common.Use(x.Position?.Value, y => $" {Common.Format(100.0 * y, configuration)}%")}"))})";
                    }
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DefaultXlsxBorderConverter"/> class.
    /// </summary>
    public class DefaultXlsxBorderConverter : IConverterBase<OpenXmlElement?, XlsxStyles>
    {
        public XlsxStyles Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxStyles result = new();

            string? style(BorderPropertiesType? border)
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
                if (style(border.TopBorder) is string top)
                {
                    result.Styles["border-top"] = top;
                }
                if (style(border.RightBorder) is string right)
                {
                    result.Styles["border-right"] = right;
                }
                if (style(border.BottomBorder) is string bottom)
                {
                    result.Styles["border-bottom"] = bottom;
                }
                if (style(border.LeftBorder) is string left)
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
    public class DefaultXlsxAlignmentConverter : IConverterBase<OpenXmlElement?, XlsxStyles>
    {
        public XlsxStyles Convert(OpenXmlElement? value, ConverterContext context, ConverterConfiguration configuration)
        {
            if (value == null)
            {
                return new();
            }

            XlsxStyles result = new();

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
                    //TODO: indent
                }
                if (alignment.WrapText != null && (alignment.WrapText?.Value ?? true))
                {
                    result.Styles["overflow-wrap"] = "break-word";
                    result.Styles["white-space"] = "normal";
                }
                if (alignment.TextRotation?.Value != null && alignment.TextRotation.Value != 0)
                {
                    result.Containers.Add(alignment.TextRotation.Value == 255 ? new()
                    {
                        ["display"] = "inline-block",
                        ["writing-mode"] = "vertical-rl",
                        ["text-orientation"] = "upright"
                    } : new()
                    {
                        ["display"] = "inline-block",
                        ["rotate"] = $"{Common.Format(alignment.TextRotation.Value switch
                        {
                            <= 90 => 360 - alignment.TextRotation.Value,
                            _ => alignment.TextRotation.Value - 90
                        }, configuration)}deg"
                    });
                }
                if (alignment.ReadingOrder?.Value != null)
                {
                    result.Styles["direction"] = alignment.ReadingOrder.Value switch
                    {
                        1 => "ltr",
                        2 => "rtl",
                        _ => "auto"
                    };
                }
            }

            return result;
        }
    }
}
