using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxToHtmlConverter.Base
{
    /// <summary>
    /// Represents an internal converter.
    /// </summary>
    /// <typeparam name="T1">The type of the source value.</typeparam>
    /// <typeparam name="T2">The type of the target value.</typeparam>
    public interface IConverterBase<T1, T2>
    {
        /// <summary>
        /// Converts the source value to the target value.
        /// </summary>
        /// <param name="value">The source value.</param>
        /// <param name="context">The conversion context.</param>
        /// <param name="configuration">The conversion configuration.</param>
        /// <returns>The conversion result.</returns>
        public abstract T2 Convert(T1 value, ConverterContext context, ConverterConfiguration configuration);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ConverterContext"/> class.
    /// </summary>
    public class ConverterContext()
    {
        /// <summary>
        /// Gets or sets the conversion cache of the context.
        /// </summary>
        public Dictionary<object, object?> Cache { get; set; } = [];

        /// <summary>
        /// Gets or sets the XLSX theme of the context.
        /// </summary>
        public DocumentFormat.OpenXml.Drawing.Theme? Theme { get; set; } = null;

        /// <summary>
        /// Gets or sets the XLSX stylesheet of the context.
        /// </summary>
        public XlsxStylesheetCollection Stylesheet { get; set; } = new();

        /// <summary>
        /// Gets or sets the XLSX shared-string table of the context.
        /// </summary>
        public XlsxString[] SharedStrings { get; set; } = [];

        /// <summary>
        /// Gets or sets the XLSX sheet of the context.
        /// </summary>
        public XlsxWorksheet Worksheet { get; set; } = new();
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlElement"/> class.
    /// </summary>
    /// <param name="indent">The level of indentation.</param>
    /// <param name="type">The type of the element.</param>
    /// <param name="tag">The tag name of the element.</param>
    /// <param name="attributes">The attributes of the element.</param>
    /// <param name="content">The content of the element.</param>
    public class HtmlElement(int indent, HtmlElement.ElementType type, string tag, HtmlAttributeCollection? attributes = null, List<object>? content = null)
    {
        /// <summary>
        /// Specifies the type of the element.
        /// </summary>
        public enum ElementType
        {
            /// <summary>
            /// Declaration element.
            /// </summary>
            Declaration,

            /// <summary>
            /// Paired element.
            /// </summary>
            Paired,

            /// <summary>
            /// Starting tag of a paired element.
            /// </summary>
            PairedStart,

            /// <summary>
            /// Closing tag of a paired element.
            /// </summary>
            PairedEnd,

            /// <summary>
            /// Unpaired element.
            /// </summary>
            Unpaired,

            /// <summary>
            /// Comment.
            /// </summary>
            Comment
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlElement"/> class.
        /// </summary>
        /// <param name="tag">The tag name of the element.</param>
        /// <param name="attributes">The attributes of the element.</param>
        /// <param name="content">The content of the element.</param>
        public HtmlElement(string tag, HtmlAttributeCollection attributes, List<object>? content = null) : this(0, ElementType.Paired, tag, attributes, content) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlElement"/> class.
        /// </summary>
        /// <param name="tag">The tag name of the element.</param>
        /// <param name="styles">The XLSX styles to apply to the element.</param>
        /// <param name="content">The content of the element.</param>
        public HtmlElement(string tag, XlsxStyles styles, List<object>? content = null) : this(tag, new HtmlAttributeCollection()
        {
            ["style"] = styles.Styles
        })
        {
            Content.AddRange(styles.ApplyContainers(content));
        }

        /// <summary>
        /// Gets or sets the level of indentation.
        /// </summary>
        public int Indent { get; set; } = indent;

        /// <summary>
        /// Gets or sets the type of the element.
        /// </summary>
        public ElementType Type { get; set; } = type;

        /// <summary>
        /// Gets or sets the tag name of the element.
        /// </summary>
        public string Tag { get; set; } = tag;

        /// <summary>
        /// Gets or sets the attributes of the element.
        /// </summary>
        public HtmlAttributeCollection Attributes { get; set; } = attributes ?? [];

        /// <summary>
        /// Gets or set the content of the element.
        /// </summary>
        public List<object> Content { get; set; } = content ?? [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlAttributeCollection"/> class.
    /// </summary>
    public class HtmlAttributeCollection() : Dictionary<string, object?>
    {
        /// <summary>
        /// Merges a specified <see cref="HtmlAttributeCollection"/> into this instance.
        /// </summary>
        /// <param name="collection">The specified <see cref="HtmlAttributeCollection"/>.</param>
        public void Merge(HtmlAttributeCollection collection)
        {
            foreach ((string key, object? value) in collection)
            {
                this[key] = value;
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlStylesheetCollection"/> class.
    /// </summary>
    public class HtmlStylesheetCollection() : Dictionary<string, HtmlStyles>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlStylesheetCollection"/> class.
        /// </summary>
        /// <param name="raw">The raw HTML representation of the stylesheet.</param>
        public HtmlStylesheetCollection(string raw) : this()
        {
            foreach (string block in raw.Split('}'))
            {
                if (!block.Contains('{'))
                {
                    continue;
                }

                string[] selector = block.Split('{');

                HtmlStyles styles = [];
                foreach (string line in selector[^1].Split(';'))
                {
                    int index = line.IndexOf(':');
                    if (index < 0)
                    {
                        continue;
                    }

                    styles[line[..index].Trim()] = line[(index + 1)..].Trim();
                }

                this[selector[0].Trim()] = styles;
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlStyles"/> class.
    /// </summary>
    public class HtmlStyles() : Dictionary<string, string>
    {
        /// <summary>
        /// Merges a specified <see cref="HtmlStyles"/> into this instance.
        /// </summary>
        /// <param name="styles">The specified <see cref="HtmlStyles"/>.</param>
        public void Merge(HtmlStyles styles)
        {
            foreach ((string key, string value) in styles)
            {
                this[key] = value;
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxWorksheet"/> class.
    /// </summary>
    public class XlsxWorksheet()
    {
        /// <summary>
        /// Gets or sets the sheet content.
        /// </summary>
        public SheetData? Data { get; set; } = null;

        /// <summary>
        /// Gets or sets the HTML styles of the sheet title.
        /// </summary>
        public HtmlStyles TitleStyles { get; set; } = [];

        /// <summary>
        /// Gets or sets the default cell size within the sheet.
        /// </summary>
        public (double Width, double Height) DefaultCellSize { get; set; } = (0, 0);

        /// <summary>
        /// Gets or sets the collection of column widths within the sheet.
        /// </summary>
        public double?[] ColumnWidths { get; set; } = [];

        /// <summary>
        /// Gets or sets the sheet dimension.
        /// </summary>
        public XlsxRange Dimension { get; set; } = new(1, 1, 1, 1);

        /// <summary>
        /// Gets or sets the collection of specialties within the sheet.
        /// </summary>
        public List<XlsxRangeSpecialty> Specialties { get; set; } = [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxRange"/> class.
    /// </summary>
    /// <param name="left">The inclusive 1-indexed start of the columns within the range.</param>
    /// <param name="top">The inclusive 1-indexed start of the rows within the range.</param>
    /// <param name="right">The inclusive 1-indexed end of the columns within the range.</param>
    /// <param name="bottom">The inclusive 1-indexed end of the rows within the range.</param>
    public class XlsxRange(uint left, uint top, uint right, uint bottom)
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxRange"/> class.
        /// </summary>
        /// <param name="range">The XLSX reference that specifies the range.</param>
        /// <param name="dimension">The dimension of the current sheet.</param>
        public XlsxRange(string range, XlsxRange? dimension = null) : this(1, 1, 1, 1)
        {
            string[] references = range.Split(':');
            (uint left, uint top) = ParseReference(references[0], (dimension?.ColumnStart ?? 1, dimension?.RowStart ?? 1));
            (uint right, uint bottom) = ParseReference(references[^1], (dimension?.ColumnEnd ?? 1, dimension?.RowEnd ?? 1));

            ColumnStart = Math.Min(left, right);
            RowStart = Math.Min(top, bottom);
            ColumnEnd = Math.Max(left, right);
            RowEnd = Math.Max(top, bottom);
        }

        /// <summary>
        /// Gets or sets the inclusive 1-indexed start of the columns within the range.
        /// </summary>
        public uint ColumnStart { get; set; } = left;

        /// <summary>
        /// Gets or sets the inclusive 1-indexed start of the rows within the range.
        /// </summary>
        public uint RowStart { get; set; } = top;

        /// <summary>
        /// Gets or sets the inclusive 1-indexed end of the columns within the range.
        /// </summary>
        public uint ColumnEnd { get; set; } = right;

        /// <summary>
        /// Gets or sets the inclusive 1-indexed end of the rows within the range.
        /// </summary>
        public uint RowEnd { get; set; } = bottom;

        /// <summary>
        /// Gets the number of columns within the range.
        /// </summary>
        public uint ColumnCount { get => ColumnEnd - ColumnStart + 1; }

        /// <summary>
        /// Gets the number of rows within the range.
        /// </summary>
        public uint RowCount { get => RowEnd - RowStart + 1; }

        /// <summary>
        /// Determines whether the range starts at the specified column and row.
        /// </summary>
        /// <param name="column">The specified column.</param>
        /// <param name="row">The specified row.</param>
        /// <returns><see langword="true"/> if the range starts at the specified column and row; otherwise, <see langword="false"/>.</returns>
        public bool StartsAt(uint column, uint row)
        {
            return column == ColumnStart && row == RowStart;
        }

        /// <summary>
        /// Determines whether the range contains the specified column and row.
        /// </summary>
        /// <param name="column">The specified column.</param>
        /// <param name="row">The specified row.</param>
        /// <returns><see langword="true"/> if the range contains the specified column and row; otherwise, <see langword="false"/>.</returns>
        public bool Contains(uint column, uint row)
        {
            return column >= ColumnStart && column <= ColumnEnd && row >= RowStart && row <= RowEnd;
        }

        /// <summary>
        /// Converts a XLSX reference that specifies a position in terms of column and row indexes.
        /// </summary>
        /// <param name="reference">The XLSX reference that specifies the position.</param>
        /// <param name="fallback">The fallback values when an index is not present.</param>
        /// <returns>The 1-indexed position.</returns>
        public static (uint Column, uint Row) ParseReference(string reference, (uint Column, uint Row)? fallback = null)
        {
            string letters = string.Concat(reference.Where(char.IsLetter));
            string digits = string.Concat(reference.Where(char.IsDigit));

            uint? column = null;
            if (letters.Any())
            {
                column = 0;
                for (int i = letters.Length - 1, multiplier = 1; i >= 0; i--, multiplier *= 26)
                {
                    column += (uint)(multiplier * (char.ToUpperInvariant(letters[i]) - 64));
                }
                column = Math.Max(1, column.Value);
            }

            uint? row = null;
            if (digits.Any())
            {
                row = Math.Max(1, uint.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint value) ? value : 1);
            }

            return (column ?? fallback?.Column ?? 1, row ?? fallback?.Row ?? 1);
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxRangeSpecialty"/> class.
    /// </summary>
    public class XlsxRangeSpecialty()
    {
        /// <summary>
        /// Gets or sets the specialty.
        /// </summary>
        public object Specialty { get; set; } = new();

        /// <summary>
        /// Gets or sets the range of the specialty.
        /// </summary>
        public XlsxRange Range { get; set; } = new(1, 1, 1, 1);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxCell"/> class.
    /// </summary>
    /// <param name="cell">The cell.</param>
    public class XlsxCell(Cell cell)
    {
        /// <summary>
        /// Gets or sets the cell.
        /// </summary>
        public Cell Cell { get; set; } = cell;

        /// <summary>
        /// Gets or sets the number format ID of the cell.
        /// </summary>
        public uint NumberFormatId { get; set; } = 0;

        /// <summary>
        /// Gets or sets the collection of specialties associated with the cell.
        /// </summary>
        public XlsxRangeSpecialty[] Specialties { get; set; } = [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxElement"/> class.
    /// </summary>
    public class XlsxString()
    {
        /// <summary>
        /// Gets or sets the content of the string.
        /// </summary>
        public List<object> Content { get; set; } = [];

        /// <summary>
        /// Gets or sets the raw representation of the string.
        /// </summary>
        public string Raw { get; set; } = string.Empty;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxContent"/> class.
    /// </summary>
    public class XlsxContent()
    {
        /// <summary>
        /// Gets or sets the content.
        /// </summary>
        public List<object> Content { get; set; } = [];

        /// <summary>
        /// Gets or sets the styles associated with the content.
        /// </summary>
        public XlsxStyles Styles { get; set; } = new();

        /// <summary>
        /// Gets or sets the collection of differential formats associated with the content.
        /// </summary>
        public List<uint> DifferentialFormatIds { get; set; } = [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxStylesheetCollection"/> class.
    /// </summary>
    public class XlsxStylesheetCollection()
    {
        /// <summary>
        /// Gets or sets the collection of XLSX cell formats.
        /// </summary>
        public XlsxCellFormat[] CellFormats { get; set; } = [];

        /// <summary>
        /// Gets or sets the collection of XLSX differential formats.
        /// </summary>
        public XlsxDifferentialFormat[] DifferentialFormats { get; set; } = [];

        /// <summary>
        /// Gets or sets the collection of XLSX number formats.
        /// </summary>
        public Dictionary<uint, XlsxNumberFormat> NumberFormats { get; set; } = [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxStyles"/> class.
    /// </summary>
    public class XlsxStyles()
    {
        /// <summary>
        /// Gets or sets the styles.
        /// </summary>
        public HtmlStyles Styles { get; set; } = [];

        /// <summary>
        /// Gets or sets the collection of styles respective to each layer of a nested HTML container element that surrounds the cell content using the styles.
        /// </summary>
        public List<HtmlStyles> Containers { get; set; } = [];

        /// <summary>
        /// Merges a specified <see cref="XlsxStyles"/> into this instance.
        /// </summary>
        /// <param name="styles">The specified <see cref="XlsxStyles"/>.</param>
        /// <param name="isReserved">Whether the styles of the specified <see cref="XlsxStyles"/> are reserved.</param>
        public void Merge(XlsxStyles styles, bool isReserved = false)
        {
            if (isReserved)
            {
                foreach (string key in styles.Styles.Select(x => x.Key))
                {
                    Styles.Remove(key);
                }
            }
            else
            {
                Styles.Merge(styles.Styles);
            }

            Containers.AddRange(styles.Containers);
        }

        /// <summary>
        /// Applies the containers onto the specified content.
        /// </summary>
        /// <param name="content">The specified content.</param>
        /// <returns>The applied HTML content.</returns>
        public List<object> ApplyContainers(List<object>? content = null)
        {
            if (!Containers.Any())
            {
                return content ?? [];
            }

            List<object> result = [];

            HtmlElement? parent = null;
            foreach (HtmlStyles container in Containers)
            {
                HtmlElement element = new("div", new HtmlAttributeCollection()
                {
                    ["style"] = container
                });

                (parent?.Content ?? result).Add(element);
                parent = element;
            }

            parent?.Content.AddRange(content ?? []);

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxCellFormat"/> class.
    /// </summary>
    public class XlsxCellFormat()
    {
        /// <summary>
        /// Gets or sets the styles of the format.
        /// </summary>
        public XlsxStyles Styles { get; set; } = new();

        /// <summary>
        /// Gets or sets the number format of the format.
        /// </summary>
        public uint NumberFormatId { get; set; } = 0;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxDifferentialFormat"/> class.
    /// </summary>
    public class XlsxDifferentialFormat()
    {
        /// <summary>
        /// Gets or sets the font styles of the format.
        /// </summary>
        public XlsxStyles FontStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the fill styles of the format.
        /// </summary>
        public XlsxStyles FillStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the border styles of the format.
        /// </summary>
        public XlsxStyles BorderStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the alignment styles of the format.
        /// </summary>
        public XlsxStyles AlignmentStyles { get; set; } = new();

        /// <summary>
        /// Combines the styles into a single <see cref="XlsxStyles"/> instance.
        /// </summary>
        /// <returns>The combined <see cref="XlsxStyles"/>.</returns>
        public XlsxStyles Combine()
        {
            XlsxStyles result = new();

            result.Merge(FillStyles);
            result.Merge(FontStyles);
            result.Merge(BorderStyles);
            result.Merge(AlignmentStyles);

            return result;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxNumberFormat"/> class.
    /// </summary>
    /// <param name="positive">The positive section of the number format.</param>
    /// <param name="negative">The negative section of the number format.</param>
    /// <param name="zero">The zero section of the number format.</param>
    /// <param name="text">The text section of the number format.</param>
    public class XlsxNumberFormat((string Code, bool IsDate)? positive, (string Code, bool IsDate)? negative, (string Code, bool IsDate)? zero, (string Code, bool IsDate)? text)
    {
        private enum EscapeState
        {
            None,
            Immediate,
            Continuous
        }

        /// <summary>
        /// Gets or sets the positive section of the number format.
        /// </summary>
        public (string Code, bool IsDate)? Positive { get; set; } = positive;

        /// <summary>
        /// Gets or sets the negative section of the number format.
        /// </summary>
        public (string Code, bool IsDate)? Negative { get; set; } = negative;

        /// <summary>
        /// Gets or sets the zero section of the number format.
        /// </summary>
        public (string Code, bool IsDate)? Zero { get; set; } = zero;

        /// <summary>
        /// Gets or sets the text section of the number format.
        /// </summary>
        public (string Code, bool IsDate)? Text { get; set; } = text;

        /// <summary>
        /// Iterates through the specified format code with regards to character escaping.
        /// </summary>
        /// <param name="code">The specified format code.</param>
        /// <param name="immediate">The additional immediate escape characters to check for.</param>
        /// <param name="continuous">The additional continuous escape characters to check for.</param>
        /// <returns>The collection of characters with information regrading character escaping.</returns>
        public static IEnumerable<(int Index, char Character, bool IsEscaped)> Escape(string code, char[]? immediate = null, char[]? continuous = null)
        {
            int index = 0;
            EscapeState state = EscapeState.None;
            foreach (char character in code)
            {
                yield return (index, character, state switch
                {
                    EscapeState.None => false,
                    EscapeState.Continuous => character is not '\"' && (!continuous?.Contains(character) ?? true),
                    _ => true
                });

                index++;
                state = (state, character) switch
                {
                    (EscapeState.None, '\\') => EscapeState.Immediate,
                    (EscapeState.None, '\"') => EscapeState.Continuous,
                    (EscapeState.None, _) when immediate?.Contains(character) ?? false => EscapeState.Immediate,
                    (EscapeState.None, _) when continuous?.Contains(character) ?? false => EscapeState.Continuous,
                    (EscapeState.Immediate, _) => EscapeState.None,
                    (EscapeState.Continuous, '\"') => EscapeState.None,
                    (EscapeState.Continuous, _) when continuous?.Contains(character) ?? false => EscapeState.None,
                    _ => state
                };
            }
        }
    }
}
