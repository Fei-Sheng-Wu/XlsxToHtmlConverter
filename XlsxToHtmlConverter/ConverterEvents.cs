using System;
using DocumentFormat.OpenXml.Packaging;

namespace XlsxToHtmlConverter
{
    /// <summary>
    /// Represents the method that will handle the progress callback event of the <see cref="Converter"/> class.
    /// </summary>
    /// <param name="sender">The source of the event.</param>
    /// <param name="e">An object that contains the event data.</param>
    public delegate void ConverterProgressChangedEventHandler(SpreadsheetDocument? sender, ConverterProgressChangedEventArgs e);

    /// <summary>
    /// Initializes a new instance of the <see cref="ConverterProgressChangedEventArgs"/> class.
    /// </summary>
    public class ConverterProgressChangedEventArgs((uint Current, uint Total) sheet, (uint Current, uint Total) row) : EventArgs
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
}
