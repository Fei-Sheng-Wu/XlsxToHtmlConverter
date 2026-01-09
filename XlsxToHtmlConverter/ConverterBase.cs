using System;
using System.Linq;
using System.Collections.Generic;
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
        public Specification.Xlsx.XlsxStylesCollection Stylesheet { get; set; } = new();

        /// <summary>
        /// Gets or sets the XLSX shared-string table of the context.
        /// </summary>
        public Specification.Xlsx.XlsxString[] SharedStrings { get; set; } = [];

        /// <summary>
        /// Gets or sets the XLSX sheet of the context.
        /// </summary>
        public Specification.Xlsx.XlsxSheet Sheet { get; set; } = new();
    }
}

namespace XlsxToHtmlConverter.Base.Specification
{
    /// <summary>
    /// Represents a mergeable object.
    /// </summary>
    public interface IMergeable
    {
        /// <summary>
        /// Merges a specified object into this instance.
        /// </summary>
        /// <param name="value">The specified object.</param>
        public abstract void Merge(object? value);
    }
}

namespace XlsxToHtmlConverter.Base.Specification.Html
{
    /// <summary>
    /// Specifies the type of a HTML element.
    /// </summary>
    public enum HtmlElementType
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
    /// <param name="indent">The indentation level of the element.</param>
    /// <param name="type">The type of the element.</param>
    /// <param name="tag">The tag name of the element.</param>
    /// <param name="attributes">The attributes of the element.</param>
    /// <param name="children">The children of the element.</param>
    public class HtmlElement(int? indent, HtmlElementType type, string tag, HtmlAttributes? attributes = null, HtmlChildren? children = null)
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlElement"/> class.
        /// </summary>
        /// <param name="type">The type of the element.</param>
        /// <param name="tag">The tag name of the element.</param>
        /// <param name="attributes">The attributes of the element.</param>
        /// <param name="children">The children of the element.</param>
        public HtmlElement(HtmlElementType type, string tag, HtmlAttributes? attributes = null, HtmlChildren? children = null) : this(null, type, tag, attributes, children) { }

        /// <summary>
        /// Gets or sets the indentation level of the element.
        /// </summary>
        public int? Indent { get; set; } = indent;

        /// <summary>
        /// Gets or sets the type of the element.
        /// </summary>
        public HtmlElementType Type { get; set; } = type;

        /// <summary>
        /// Gets or sets the tag name of the element.
        /// </summary>
        public string Tag { get; set; } = tag;

        /// <summary>
        /// Gets or sets the attributes of the element.
        /// </summary>
        public HtmlAttributes Attributes { get; set; } = attributes ?? [];

        /// <summary>
        /// Gets or set the children of the element.
        /// </summary>
        public HtmlChildren Children { get; set; } = children ?? [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlChildren"/> class.
    /// </summary>
    public class HtmlChildren() : List<object> { }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlAttributes"/> class.
    /// </summary>
    public class HtmlAttributes() : Dictionary<string, object?>, IMergeable
    {
        public void Merge(object? value)
        {
            if (value is not HtmlAttributes attributes)
            {
                return;
            }

            foreach ((string attribute, object? field) in attributes)
            {
                if (field is IMergeable mergeable && Implementation.Common.Get(this, attribute) is IMergeable baseline)
                {
                    baseline.Merge(mergeable);
                    continue;
                }

                this[attribute] = field;
            }
        }

        /// <summary>
        /// Merges a <see cref="HtmlStyles"/> instance into this instance.
        /// </summary>
        /// <param name="styles">The <see cref="HtmlStyles"/> instance.</param>
        /// <param name="name">The class name of the <see cref="HtmlStyles"/> instance.</param>
        public void MergeStyles(HtmlStyles styles, string? name = null)
        {
            if (Implementation.Common.Get(this, "style") is not HtmlStyles baseline)
            {
                baseline = [];
                this["style"] = baseline;
            }
            if (name == null)
            {
                baseline.Merge(styles);
                return;
            }

            if (Implementation.Common.Get(this, "class") is not HtmlClasses classes)
            {
                classes = [name];
                this["class"] = classes;
            }
            else
            {
                classes.Add(name);
            }

            foreach (string property in styles.Keys)
            {
                baseline.Remove(property);
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlClasses"/> class.
    /// </summary>
    public class HtmlClasses() : List<string>, IMergeable
    {
        public void Merge(object? value)
        {
            if (value is not HtmlClasses classes)
            {
                return;
            }

            AddRange(classes);
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="HtmlStylesCollection"/> class.
    /// </summary>
    public class HtmlStylesCollection() : Dictionary<string, HtmlStyles>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlStylesCollection"/> class.
        /// </summary>
        /// <param name="raw">The raw HTML representation of the stylesheet.</param>
        public HtmlStylesCollection(string raw) : this()
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
    public class HtmlStyles() : Dictionary<string, string>, IMergeable
    {
        public void Merge(object? value)
        {
            if (value is not HtmlStyles styles)
            {
                return;
            }

            foreach ((string property, string field) in styles)
            {
                this[property] = field;
            }
        }
    }
}

namespace XlsxToHtmlConverter.Base.Specification.Xlsx
{
    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxCell"/> class.
    /// </summary>
    /// <param name="cell">The cell.</param>
    public class XlsxCell(Cell? cell)
    {
        /// <summary>
        /// Gets or sets the cell.
        /// </summary>
        public Cell? Cell { get; set; } = cell;

        /// <summary>
        /// Gets or sets the styles of the cell.
        /// </summary>
        public List<XlsxStyles> Styles { get; set; } = [];

        /// <summary>
        /// Gets or sets the number format of the cell.
        /// </summary>
        public XlsxNumberFormat? NumberFormat { get; set; } = null;

        /// <summary>
        /// Gets or sets the number format ID of the cell.
        /// </summary>
        public uint? NumberFormatId { get; set; } = null;

        /// <summary>
        /// Gets or sets the specialties associated with the cell.
        /// </summary>
        public IEnumerable<XlsxSpecialty> Specialties { get; set; } = [];

        /// <summary>
        /// Gets or sets the attributes of the cell.
        /// </summary>
        public Html.HtmlAttributes Attributes { get; set; } = [];

        /// <summary>
        /// Gets or sets the children of the cell.
        /// </summary>
        public Html.HtmlChildren Children { get; set; } = [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxString"/> class.
    /// </summary>
    public class XlsxString()
    {
        /// <summary>
        /// Gets or sets the children of the string.
        /// </summary>
        public Html.HtmlChildren Children { get; set; } = [];

        /// <summary>
        /// Gets or sets the raw representation of the string.
        /// </summary>
        public string Raw { get; set; } = string.Empty;
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
        public XlsxRange() : this(1, 1, 1, 1) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxRange"/> class.
        /// </summary>
        /// <param name="raw">The XLSX reference that specifies the range.</param>
        /// <param name="dimension">The dimension of the current sheet.</param>
        public XlsxRange(string raw, XlsxRange? dimension = null) : this()
        {
            string[] references = raw.Split(':');
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
            return ContainsColumn(column) && ContainsRow(row);
        }

        /// <summary>
        /// Determines whether the range contains the specified column.
        /// </summary>
        /// <param name="column">The specified column.</param>
        /// <returns><see langword="true"/> if the range contains the specified column; otherwise, <see langword="false"/>.</returns>
        public bool ContainsColumn(uint column)
        {
            return column >= ColumnStart && column <= ColumnEnd;
        }

        /// <summary>
        /// Determines whether the range contains the specified row.
        /// </summary>
        /// <param name="row">The specified row.</param>
        /// <returns><see langword="true"/> if the range contains the specified row; otherwise, <see langword="false"/>.</returns>
        public bool ContainsRow(uint row)
        {
            return row >= RowStart && row <= RowEnd;
        }

        /// <summary>
        /// Converts a XLSX reference that represents a specified column and row.
        /// </summary>
        /// <param name="reference">The XLSX reference.</param>
        /// <param name="fallback">The fallback values when an index is not present.</param>
        /// <returns>The 1-indexed specified column and row.</returns>
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
                row = Math.Max(1, Implementation.Common.ParsePositive(digits) ?? 1);
            }

            return (column ?? fallback?.Column ?? 1, row ?? fallback?.Row ?? 1);
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxSheet"/> class.
    /// </summary>
    public class XlsxSheet()
    {
        /// <summary>
        /// Gets or sets the sheet content.
        /// </summary>
        public SheetData? Data { get; set; } = null;

        /// <summary>
        /// Gets or sets the sheet dimension.
        /// </summary>
        public XlsxRange Dimension { get; set; } = new();

        /// <summary>
        /// Gets or sets the attributes of the sheet title.
        /// </summary>
        public Html.HtmlAttributes TitleAttributes { get; set; } = [];

        /// <summary>
        /// Gets or sets the attributes of the columns within the sheet.
        /// </summary>
        public Html.HtmlAttributes ColumnAttributes { get; set; } = [];

        /// <summary>
        /// Gets or sets the attributes of the rows within the sheet.
        /// </summary>
        public Html.HtmlAttributes RowAttributes { get; set; } = [];

        /// <summary>
        /// Gets or sets the attributes of the cells within the sheet.
        /// </summary>
        public Html.HtmlAttributes CellAttributes { get; set; } = [];

        /// <summary>
        /// Gets or sets the columns within the sheet.
        /// </summary>
        public (double? Width, bool? IsHidden, uint? StylesIndex)[] Columns { get; set; } = [];

        /// <summary>
        /// Gets or sets the size of the cells within the sheet.
        /// </summary>
        public (double Width, double Height) CellSize { get; set; } = (0, 0);

        /// <summary>
        /// Gets or sets the specialties within the sheet.
        /// </summary>
        public List<XlsxSpecialty> Specialties { get; set; } = [];
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxSpecialty"/> class.
    /// </summary>
    /// <param name="specialty">The specialty.</param>
    public class XlsxSpecialty(object specialty)
    {
        /// <summary>
        /// Gets or sets the specialty.
        /// </summary>
        public object Specialty { get; set; } = specialty;

        /// <summary>
        /// Gets or sets the range of the specialty.
        /// </summary>
        public XlsxRange Range { get; set; } = new();
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxStylesCollection"/> class.
    /// </summary>
    public class XlsxStylesCollection()
    {
        /// <summary>
        /// Gets or sets the base styles.
        /// </summary>
        public XlsxBaseStyles[] BaseStyles { get; set; } = [];

        /// <summary>
        /// Gets or sets the differential styles.
        /// </summary>
        public XlsxDifferentialStyles[] DifferentialStyles { get; set; } = [];

        /// <summary>
        /// Gets or sets the number formats.
        /// </summary>
        public Dictionary<uint, XlsxNumberFormat> NumberFormats { get; set; } = [];
    }

    /// <summary>
    /// Provides applicable XLSX styles.
    /// </summary>
    public abstract class XlsxStyles
    {
        /// <summary>
        /// Gets or sets the name of the styles.
        /// </summary>
        public string? Name { get; set; } = null;

        /// <summary>
        /// Gets or sets whether the styles should hide the cell content.
        /// </summary>
        public bool? IsHidden { get; set; } = null;

        protected abstract IEnumerable<XlsxStylesLayer> GetLayers();

        /// <summary>
        /// Retrieves the styles.
        /// </summary>
        public Html.HtmlStyles GetsStyles()
        {
            return GetsStyles(GetLayers());
        }

        /// <summary>
        /// Applies the styles onto the specified HTML element.
        /// </summary>
        /// <param name="element">The specified HTML element.</param>
        /// <param name="isNamed">Whether to use the names of the styles.</param>
        public void ApplyStyles(Html.HtmlElement element, bool isNamed)
        {
            ApplyStyles(element, GetLayers(), isNamed ? Name : null);
        }

        /// <summary>
        /// Retrieves the styles.
        /// </summary>
        /// <param name="layers">The layers of the styles.</param>
        public static Html.HtmlStyles GetsStyles(IEnumerable<XlsxStylesLayer> layers)
        {
            Html.HtmlStyles result = [];

            foreach (XlsxStylesLayer layer in layers)
            {
                result.Merge(layer.Styles);
            }

            return result;
        }

        /// <summary>
        /// Applies the styles onto the specified HTML element.
        /// </summary>
        /// <param name="element">The specified HTML element.</param>
        /// <param name="layers">The layers of the styles.</param>
        /// <param name="name">The name of the styles.</param>
        public static void ApplyStyles(Html.HtmlElement element, IEnumerable<XlsxStylesLayer> layers, string? name = null)
        {
            element.Attributes.MergeStyles(GetsStyles(layers), name);

            foreach (XlsxStylesLayer layer in layers)
            {
                Html.HtmlChildren children = element.Children;
                foreach (Html.HtmlStyles container in layer.Containers)
                {
                    children = [new Html.HtmlElement(Html.HtmlElementType.Paired, "span", new Html.HtmlAttributes()
                    {
                        ["style"] = container
                    }, children)];
                }

                element.Children = children;
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxStylesLayer"/> class.
    /// </summary>
    public class XlsxStylesLayer : IMergeable
    {
        /// <summary>
        /// Gets or sets the styles of the styles layer.
        /// </summary>
        public Html.HtmlStyles Styles { get; set; } = [];

        /// <summary>
        /// Gets or sets the styles respective to each layer of nested containers that surround the cell content using the styles layer.
        /// </summary>
        public List<Html.HtmlStyles> Containers { get; set; } = [];

        public void Merge(object? value)
        {
            if (value is not XlsxStylesLayer layer)
            {
                return;
            }

            Styles.Merge(layer.Styles);
            Containers.AddRange(layer.Containers);
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxBaseStyles"/> class.
    /// </summary>
    public class XlsxBaseStyles() : XlsxStyles
    {
        /// <summary>
        /// Gets or sets the styles.
        /// </summary>
        public XlsxStylesLayer Styles { get; set; } = new();

        /// <summary>
        /// Gets or sets the number format ID of the styles.
        /// </summary>
        public uint? NumberFormatId { get; set; } = null;

        protected override IEnumerable<XlsxStylesLayer> GetLayers()
        {
            yield return Styles;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxDifferentialStyles"/> class.
    /// </summary>
    public class XlsxDifferentialStyles() : XlsxStyles
    {
        /// <summary>
        /// Gets or sets the font styles of the styles.
        /// </summary>
        public XlsxStylesLayer FontStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the fill styles of the styles.
        /// </summary>
        public XlsxStylesLayer FillStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the border styles of the styles.
        /// </summary>
        public XlsxStylesLayer BorderStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the alignment styles of the styles.
        /// </summary>
        public XlsxStylesLayer AlignmentStyles { get; set; } = new();

        /// <summary>
        /// Gets or sets the number format of the styles.
        /// </summary>
        public XlsxNumberFormat? NumberFormat { get; set; } = null;

        protected override IEnumerable<XlsxStylesLayer> GetLayers()
        {
            yield return FontStyles;
            yield return FillStyles;
            yield return BorderStyles;
            yield return AlignmentStyles;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxNumberFormat"/> class.
    /// </summary>
    /// <param name="positive">The positive section of the number format.</param>
    /// <param name="negative">The negative section of the number format.</param>
    /// <param name="zero">The zero section of the number format.</param>
    /// <param name="text">The text section of the number format.</param>
    public class XlsxNumberFormat(XlsxNumberFormatCode positive, XlsxNumberFormatCode negative, XlsxNumberFormatCode zero, XlsxNumberFormatCode text)
    {
        protected enum EscapeState
        {
            None,
            Immediate,
            Continuous
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxNumberFormat"/> class.
        /// </summary>
        /// <param name="positive">The positive section of the number format.</param>
        /// <param name="negative">The negative section of the number format.</param>
        /// <param name="zero">The zero section of the number format.</param>
        public XlsxNumberFormat(XlsxNumberFormatCode positive, XlsxNumberFormatCode negative, XlsxNumberFormatCode zero) : this(positive, negative, zero, new()) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxNumberFormat"/> class.
        /// </summary>
        /// <param name="positive">The positive section of the number format.</param>
        /// <param name="negative">The negative section of the number format.</param>
        public XlsxNumberFormat(XlsxNumberFormatCode positive, XlsxNumberFormatCode negative) : this(positive, negative, positive) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxNumberFormat"/> class.
        /// </summary>
        /// <param name="positive">The positive section of the number format.</param>
        public XlsxNumberFormat(XlsxNumberFormatCode positive) : this(positive, new($"-{positive.Code}", positive.IsDate)) { }

        /// <summary>
        /// Gets or sets the positive section of the number format.
        /// </summary>
        public XlsxNumberFormatCode Positive { get; set; } = positive;

        /// <summary>
        /// Gets or sets the negative section of the number format.
        /// </summary>
        public XlsxNumberFormatCode Negative { get; set; } = negative;

        /// <summary>
        /// Gets or sets the zero section of the number format.
        /// </summary>
        public XlsxNumberFormatCode Zero { get; set; } = zero;

        /// <summary>
        /// Gets or sets the text section of the number format.
        /// </summary>
        public XlsxNumberFormatCode Text { get; set; } = text;

        /// <summary>
        /// Iterates through the specified format code with regards to character escaping.
        /// </summary>
        /// <param name="code">The specified format code.</param>
        /// <param name="singles">The additional immediate escape characters to check for.</param>
        /// <param name="blocks">The additional continuous escape characters to check for.</param>
        /// <returns>The characters with respective information regrading character escaping.</returns>
        public static IEnumerable<(int Index, char Character, bool IsEscaped)> Escape(string code, char[]? singles = null, char[]? blocks = null)
        {
            int index = 0;
            EscapeState state = EscapeState.None;

            foreach (char character in code)
            {
                yield return (index, character, state switch
                {
                    EscapeState.None => false,
                    EscapeState.Continuous => character is not '\"' && (!blocks?.Contains(character) ?? true),
                    _ => true
                });

                index++;
                state = (state, character) switch
                {
                    (EscapeState.None, '\\') => EscapeState.Immediate,
                    (EscapeState.None, '\"') => EscapeState.Continuous,
                    (EscapeState.None, _) when singles?.Contains(character) ?? false => EscapeState.Immediate,
                    (EscapeState.None, _) when blocks?.Contains(character) ?? false => EscapeState.Continuous,
                    (EscapeState.Immediate, _) => EscapeState.None,
                    (EscapeState.Continuous, '\"') => EscapeState.None,
                    (EscapeState.Continuous, _) when blocks?.Contains(character) ?? false => EscapeState.None,
                    _ => state
                };
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxNumberFormatCode"/> class.
    /// </summary>
    /// <param name="code">The code.</param>
    /// <param name="isDate">Whether the code is a date representation.</param>
    public class XlsxNumberFormatCode(string code, bool isDate)
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxNumberFormatCode"/> class.
        /// </summary>
        public XlsxNumberFormatCode() : this(string.Empty, false) { }

        /// <summary>
        /// Gets or sets the code.
        /// </summary>
        public string Code { get; set; } = code;

        /// <summary>
        /// Gets or sets whether the code is a date representation.
        /// </summary>
        public bool IsDate { get; set; } = isDate;
    }
}
