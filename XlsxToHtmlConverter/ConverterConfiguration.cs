using System;
using System.Text;
using System.Collections.Generic;
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
            public Base.IConverterBase<Base.Specification.Html.HtmlElement, string> HtmlWriter { get; set; } = new Base.Implementation.DefaultHtmlWriter();

            /// <summary>
            /// Gets or sets the converter to iterate the XLSX sheets.
            /// </summary>
            public Base.IConverterBase<Base.Specification.Xlsx.XlsxSheet?, IEnumerable<Base.Specification.Xlsx.XlsxCell>> XlsxWorksheetIterator { get; set; } = new Base.Implementation.DefaultXlsxWorksheetIterator();

            /// <summary>
            /// Gets or sets the converter to read the XLSX stylesheet.
            /// </summary>
            public Base.IConverterBase<Stylesheet?, Base.Specification.Xlsx.XlsxStylesCollection> XlsxStylesheetReader { get; set; } = new Base.Implementation.DefaultXlsxStylesheetReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX shared-string table.
            /// </summary>
            public Base.IConverterBase<SharedStringTable?, Base.Specification.Xlsx.XlsxString[]> XlsxSharedStringTableReader { get; set; } = new Base.Implementation.DefaultXlsxSharedStringTableReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX sheets.
            /// </summary>
            public Base.IConverterBase<Worksheet?, Base.Specification.Xlsx.XlsxSheet> XlsxWorksheetReader { get; set; } = new Base.Implementation.DefaultXlsxWorksheetReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX cell values.
            /// </summary>
            public Base.IConverterBase<Base.Specification.Xlsx.XlsxCell?, Base.Specification.Xlsx.XlsxCell> XlsxCellContentReader { get; set; } = new Base.Implementation.DefaultXlsxCellContentReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX tables.
            /// </summary>
            public Base.IConverterBase<TableDefinitionPart?, IEnumerable<Base.Specification.Xlsx.XlsxSpecialty>> XlsxTableReader { get; set; } = new Base.Implementation.DefaultXlsxTableReader();

            /// <summary>
            /// Gets or sets the converter to read the XLSX drawings.
            /// </summary>
            public Base.IConverterBase<DrawingsPart?, IEnumerable<Base.Specification.Xlsx.XlsxSpecialty>> XlsxDrawingReader { get; set; } = new Base.Implementation.DefaultXlsxDrawingReader();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX colors.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, string> XlsxColorConverter { get; set; } = new Base.Implementation.DefaultXlsxColorConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX strings.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.Specification.Xlsx.XlsxString> XlsxStringConverter { get; set; } = new Base.Implementation.DefaultXlsxStringConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX fonts.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.Specification.Xlsx.XlsxStylesLayer> XlsxFontConverter { get; set; } = new Base.Implementation.DefaultXlsxFontConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX fills.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.Specification.Xlsx.XlsxStylesLayer> XlsxFillConverter { get; set; } = new Base.Implementation.DefaultXlsxFillConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX borders.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.Specification.Xlsx.XlsxStylesLayer> XlsxBorderConverter { get; set; } = new Base.Implementation.DefaultXlsxBorderConverter();

            /// <summary>
            /// Gets or sets the converter to convert the XLSX alignments.
            /// </summary>
            public Base.IConverterBase<OpenXmlElement?, Base.Specification.Xlsx.XlsxStylesLayer> XlsxAlignmentConverter { get; set; } = new Base.Implementation.DefaultXlsxAlignmentConverter();
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
        /// Gets or sets the tab character for writing the HTML content.
        /// </summary>
        public string TabCharacter { get; set; } = new(' ', 2);

        /// <summary>
        /// Gets or sets the number of decimal places for rounding numeric HTML attributes. If set to negative values, no rounding is used.
        /// </summary>
        public int RoundingDigits { get; set; } = 2;

        /// <summary>
        /// Gets or sets the local culture for converting culture-dependent XLSX content.
        /// </summary>
        public CultureInfo CurrentCulture { get; set; } = CultureInfo.CurrentCulture;

        /// <summary>
        /// Gets or sets the title of the HTML content.
        /// </summary>
        public string? HtmlTitle { get; set; } = null;

        /// <summary>
        /// Gets or sets the root class of the HTML content.
        /// </summary>
        public string? HtmlRootClass { get; set; } = "xlsx";

        /// <summary>
        /// Gets or sets the preset stylesheet of the HTML content.
        /// </summary>
        public Base.Specification.Html.HtmlStylesCollection HtmlPresetStylesheet { get; set; } = new()
        {
            ["table"] = new()
            {
                ["table-layout"] = "fixed",
                ["border-collapse"] = "collapse"
            },
            ["caption"] = new()
            {
                ["margin"] = "10px auto",
                ["padding"] = "2px",
                ["width"] = "fit-content",
                ["font-size"] = "20px",
                ["font-weight"] = "bold",
                ["border-bottom"] = "thick solid var(--sheet-color)"
            },
            ["td"] = new()
            {
                ["padding"] = "0 2px",
                ["vertical-align"] = "bottom",
                ["line-height"] = "1.25",
                ["border"] = "thin solid lightgray",
                ["white-space"] = "preserve nowrap",
                ["overflow-y"] = "clip",
                ["box-sizing"] = "border-box"
            }
        };

        /// <summary>
        /// Gets or sets the selector that determines whether a XLSX sheet should be converted.
        /// </summary>
        public Func<(int Index, string? Id), bool>? XlsxSheetSelector { get; set; } = null;

        /// <summary>
        /// Gets or sets the selector that determines the dimension of a XLSX sheet.
        /// </summary>
        public Func<(uint Left, uint Top, uint Right, uint Bottom), (uint Left, uint Top, uint Right, uint Bottom)>? XlsxSheetDimensionSelector { get; set; } = null;
        
        /// <summary>
        /// Gets or sets the selector that determines whether a XLSX cell should be converted.
        /// </summary>
        public Func<(uint Column, uint Row), bool>? XlsxCellSelector { get; set; } = null;

        /// <summary>
        /// Gets or sets the selector that determines whether a XLSX object should be converted.
        /// </summary>
        public Func<((uint Column, uint Row)? Start, (uint Column, uint Row)? End), bool>? XlsxObjectSelector { get; set; } = null;

        /// <summary>
        /// Gets or sets whether to convert XLSX sheet names to HTML table captions.
        /// </summary>
        public bool ConvertSheetTitles { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX cell sizes.
        /// </summary>
        public bool ConvertSizes { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX visibilities.
        /// </summary>
        public bool ConvertVisibilities { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to convert XLSX visual styles.
        /// </summary>
        public bool ConvertStyles { get; set; } = true;

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

        /// <summary>
        /// Gets or sets whether to utilize the HTML proportional widths with percentages.
        /// </summary>
        public bool UseHtmlProportionalWidths { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to utilize the HTML machine-readable elements.
        /// </summary>
        public bool UseHtmlDataElements { get; set; } = true;
    }
}
