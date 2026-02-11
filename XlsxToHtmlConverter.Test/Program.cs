using System;
using System.IO;
using System.Linq;
using System.Diagnostics;

namespace XlsxToHtmlConverter.Test
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string? xlsx = args.Length > 0 ? args[0].Trim('\"') : null;
            string? html = args.Length > 1 ? args[1].Trim('\"') : Path.ChangeExtension(xlsx, "html");

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
                Console.Write($"{e.ProgressPercentage:F2}% (Sheet {e.CurrentSheet} of {e.SheetCount} | Row {e.CurrentRow} of {e.RowCount})    {new string('█', (int)Math.Round(e.ProgressPercentage / 100.0 * 50)).PadRight(50, '░')}");
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
