using System.Text;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxToHtmlConverter
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ConverterConfiguration"/> class.
    /// </summary>
    public class ConverterConfiguration()
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Composition"/> class.
        /// </summary>
        public class Composition()
        {
            /// <summary>
            /// Gets or sets the converter to write the HTML content.
            /// </summary>
            public Base.IConverterBase<Base.HtmlElement, string> HtmlWriter { get; set; } = new Base.Defaults.DefaultHtmlWriter();

            /// <summary>
            /// Gets or sets the converter to read the XLSX stylesheet.
            /// </summary>
            public Base.IConverterBase<Stylesheet?, Base.XlsxStylesheetCollection> XlsxStylesheetReader { get; set; } = new Base.Defaults.DefaultXlsxStylesheetReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX shared-string table.
            /// </summary>
            public Base.IConverterBase<SharedStringTable?, Base.XlsxString[]> XlsxSharedStringTableReader { get; set; } = new Base.Defaults.DefaultXlsxSharedStringTableReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX sheets.
            /// </summary>
            public Base.IConverterBase<Worksheet?, Base.XlsxWorksheet> XlsxWorksheetReader { get; set; } = new Base.Defaults.DefaultXlsxWorksheetReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX cell values.
            /// </summary>
            public Base.IConverterBase<Base.XlsxCell?, Base.XlsxContent> XlsxCellContentReader { get; set; } = new Base.Defaults.DefaultXlsxCellContentReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX tables.
            /// </summary>
            public Base.IConverterBase<TableDefinitionPart?, Base.XlsxRangeSpecialty[]> XlsxTableReader { get; set; } = new Base.Defaults.DefaultXlsxTableReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX drawings.
            /// </summary>
            public Base.IConverterBase<DrawingsPart?, Base.XlsxRangeSpecialty[]> XlsxDrawingReader { get; set; } = new Base.Defaults.DefaultXlsxDrawingReader();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX colors.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, string> XlsxColorConverter { get; set; } = new Base.Defaults.DefaultXlsxColorConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX strings.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.XlsxString> XlsxStringConverter { get; set; } = new Base.Defaults.DefaultXlsxStringConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX fonts.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.XlsxStyles> XlsxFontConverter { get; set; } = new Base.Defaults.DefaultXlsxFontConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX fills.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.XlsxStyles> XlsxFillConverter { get; set; } = new Base.Defaults.DefaultXlsxFillConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX borders.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.XlsxStyles> XlsxBorderConverter { get; set; } = new Base.Defaults.DefaultXlsxBorderConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX alignments.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.XlsxStyles> XlsxAlignmentConverter { get; set; } = new Base.Defaults.DefaultXlsxAlignmentConverter();
        }

        /// <summary>
        /// Gets or sets the internal converters to use during the conversion process.
        /// </summary>
        public Composition ConverterComposition { get; set; } = new();

        /// <summary>
        /// Gets or sets the buffer size for writing the HTML content.
        /// </summary>
        public int BufferSize { get; set; } = 65536;

        /// <summary>
        /// Gets or sets the encoding for writing the HTML content.
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;

        /// <summary>
        /// Gets or sets the newline character for writing the HTML content.
        /// </summary>
        public string NewlineCharacter { get; set; } = "\n";

        /// <summary>
        /// Gets or sets the number of decimal places for rounding numeric HTML attributes. If set to negative values, no rounding is used.
        /// </summary>
        public int RoundingDigits { get; set; } = 2;

        public CultureInfo CurrentCulture { get; set; } = CultureInfo.CurrentCulture;

        /// <summary>
        /// Gets or sets the title of the HTML content.
        /// </summary>
        public string HtmlTitle { get; set; } = "Conversion Result";

        /// <summary>
        /// Gets or sets the preset stylesheet of the HTML content.
        /// </summary>
        public Base.HtmlStylesheetCollection HtmlPresetStylesheet { get; set; } = new()
        {
            ["body"] = new()
            {
                { "margin", "0" },
                { "padding", "0" }
            },
            ["table"] = new()
            {
                { "width", "100%" },
                { "table-layout", "fixed" },
                { "border-collapse", "collapse" }
            },
            ["caption"] = new()
            {
                { "margin", "10px auto" },
                { "padding", "2px" },
                { "width", "fit-content" },
                { "font-size", "20px" },
                { "font-weight", "bold" }
            },
            ["td"] = new()
            {
                { "padding", "0" },
                { "text-align", "left" },
                { "vertical-align", "bottom" },
                { "border", "thin solid lightgray" },
                { "white-space", "pre" },
                { "overflow", "hidden" },
                { "box-sizing", "border-box" }
            }
        };

        /// <summary>
        /// Gets or sets whether to convert XLSX sheet names to HTML table captions.
        /// </summary>
        public bool ConvertSheetTitles { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX hidden sheets.
        /// </summary>
        public bool ConvertHiddenSheets { get; set; } = false;

        /// <summary>
        /// Gets or sets whether to convert the first XLSX sheet only.
        /// </summary>
        public bool ConvertFirstSheetOnly { get; set; } = false;

        /// <summary>
        /// Gets or sets whether to convert XLSX visual styles.
        /// </summary>
        public bool ConvertStyles { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX cell sizes.
        /// </summary>
        public bool ConvertSizes { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX cell values using the respective number formats.
        /// </summary>
        public bool ConvertNumberFormats { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX pictures to HTML images.
        /// </summary>
        public bool ConvertPictures { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX shapes to HTML elements.
        /// </summary>
        public bool ConvertShapes { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert the XLSX content to a HTML fragment without the head elements.
        /// </summary>
        public bool UseHtmlFragment { get; set; } = false;

        /// <summary>
        /// Gets or sets whether to utilize the HTML class attribute.
        /// </summary>
        public bool UseHtmlClasses { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to utilize the HTML hexadecimal color representations.
        /// </summary>
        public bool UseHtmlHexColors { get; set; } = true;
    }
}
