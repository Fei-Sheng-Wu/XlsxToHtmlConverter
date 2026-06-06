using System;
using System.IO;
using System.Linq;
using System.Diagnostics;

namespace XlsxToHtmlConverter.Test
{
    public class Program
    {
        public static void Main(string[] arguments)
        {
            string? xlsx = arguments.Length > 0 ? arguments[0].Trim('\"') : null;
            string? html = arguments.Length > 1 ? arguments[1].Trim('\"') : Path.ChangeExtension(xlsx, "html");

            while (xlsx == null || !File.Exists(xlsx))
            {
                Console.WriteLine("Please enter the path to the input XLSX file:");
                xlsx = Console.ReadLine()?.Trim('\"');
            }
            if (html == null)
            {
                Console.WriteLine("Please enter the path to the output HTML file:");
                html = Console.ReadLine()?.Trim('\"') is string output && output.Any() ? output : Path.ChangeExtension(xlsx, "html");
            }

            Console.WriteLine();

            Stopwatch stopwatch = Stopwatch.StartNew();

            Converter.Convert(xlsx, html, new()
            {
                HtmlTitle = Path.GetFileName(xlsx)
            }, (x, e) =>
            {
                Console.SetCursorPosition(0, Console.CursorTop);
                Console.Write($"{e.ProgressPercentage:F2}% (Sheet {e.CurrentSheet} of {e.SheetCount} | Row {e.CurrentRow} of {e.RowCount})    {new string('█', (int)Math.Round(50.0 * e.ProgressPercentage / 100.0)).PadRight(50, '░')}");
            });

            Console.WriteLine();
            Console.WriteLine($"The conversion finished after {stopwatch.Elapsed}.");

            Console.WriteLine();
            Console.WriteLine("Press Enter to open the HTML file.");
            Console.ReadLine();

            using Process? process = Process.Start(new ProcessStartInfo()
            {
                FileName = html,
                UseShellExecute = true
            });
        }
    }
}
