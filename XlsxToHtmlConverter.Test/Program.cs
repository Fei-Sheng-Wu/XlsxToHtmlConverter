using System;
using System.IO;

namespace XlsxToHtmlConverter.Test
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.Clear();
            Console.SetCursorPosition(0, 0);
            Console.CursorVisible = false;

            string xlsxFilePath;
            string htmlFilePath;

            //Get the input and output file paths
            if (args != null && args.Length == 2)
            {
                xlsxFilePath = args[0];
                htmlFilePath = args[1];
            }
            else if (args != null && args.Length == 1)
            {
                xlsxFilePath = args[0];
                htmlFilePath = Path.ChangeExtension(xlsxFilePath, "html");
            }
            else
            {
                Console.WriteLine("Please enter the path to the Xlsx file:");

                Console.CursorVisible = true;
                //xlsxFilePath = Console.ReadLine();
                xlsxFilePath = "E:\\Personal\\SVN\\C# Library\\XlsxToHtmlConverter\\sample.xlsx";
                Console.CursorVisible = false;

                Console.WriteLine("Please enter the path to the Html file:");

                Console.CursorVisible = true;
                //htmlFilePath = Console.ReadLine();
                htmlFilePath = "E:\\Personal\\SVN\\C# Library\\XlsxToHtmlConverter\\sample.html";
                Console.CursorVisible = false;
            }

            Console.WriteLine();
            Console.WriteLine();

            try
            {
                //Set up the progress callback
                EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> progressCallback = ConverterProgressCallback;

                //Adjust the conversion configurations
                XlsxToHtmlConverter.ConverterConfig config = new XlsxToHtmlConverter.ConverterConfig()
                {
                    PageTitle = Path.GetFileName(xlsxFilePath)
                };

                //Convert the Xlsx file
                using (FileStream outputStream = new FileStream(htmlFilePath, FileMode.Create))
                {
                    int time = Environment.TickCount;
                    XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFilePath, outputStream, config, progressCallback);
                    Console.WriteLine($"\nThe conversion costed {TimeSpan.FromMilliseconds(Environment.TickCount - time).TotalSeconds} seconds.");
                }

                //Open the Html file
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(htmlFilePath) { UseShellExecute = true, CreateNoWindow = true });
                }
                finally
                {
                    Console.WriteLine("\n\nPress Enter to exit.");
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                //Output the error
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine($"\nError: {ex.Message}");
                Console.WriteLine("\n\nPress Enter to exit.");
                Console.ReadLine();
            }
        }

        private static void ConverterProgressCallback(object sender, XlsxToHtmlConverter.ConverterProgressCallbackEventArgs e)
        {
            //Output the progress
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.WriteLine($"{e.ProgressPercent:##0.00}% (Sheet {e.CurrentSheet} of {e.TotalSheets} | Row {e.CurrentRow} of {e.TotalRows}){new string(' ', 5) + new string('█', (int)(e.ProgressPercent / 2)) + new string('░', (int)((100 - e.ProgressPercent) / 2))}");
        }
    }
}
