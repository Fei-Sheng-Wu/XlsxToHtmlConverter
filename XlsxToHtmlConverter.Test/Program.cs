using System;
using System.IO;
using System.Diagnostics;

namespace XlsxToHtmlConverter.Test
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string? xlsx = args.Length > 0 ? args[0] : null;
            string? html = args.Length > 1 ? args[1] : Path.ChangeExtension(xlsx, "html");

            while (xlsx == null || !File.Exists(xlsx))
            {
                Console.WriteLine("Please enter the path to the input XLSX file:");
                xlsx = Console.ReadLine();
            }
            while (html == null)
            {
                Console.WriteLine("Please enter the path to the output HTML file:");
                html = Console.ReadLine();
            }

            Console.WriteLine();

            Stopwatch stopwatch = Stopwatch.StartNew();

            Converter.ConvertXlsx(xlsx, html, new ConverterConfiguration()
            {
                HtmlTitle = Path.GetFileName(xlsx)
            }, (x, e) =>
            {
                Console.SetCursorPosition(0, Console.CursorTop);
                Console.Write($"{e.ProgressPercentage:F2}% (Sheet {e.CurrentSheet} of {e.SheetCount} | Row {e.CurrentRow} of {e.RowCount})    {new string('█', (int)(e.ProgressPercentage / 2)).PadRight(50, '░')}");
            });

            Console.WriteLine();
            Console.WriteLine($"The conversion finished after {stopwatch.Elapsed}.");
            Console.WriteLine();
            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }
    }
}
