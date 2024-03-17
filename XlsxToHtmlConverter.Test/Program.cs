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

            string xlsxFileName;
            string htmlFileName;

            //Get the input and output file paths
            if (args != null && args.Length == 2)
            {
                xlsxFileName = args[0];
                htmlFileName = args[1];
            }
            else if (args != null && args.Length == 1)
            {
                xlsxFileName = args[0];
                htmlFileName = Path.ChangeExtension(xlsxFileName, "html");
            }
            else
            {
                Console.WriteLine("Please enter xlsx file path:");

                Console.CursorVisible = true;
                xlsxFileName = Console.ReadLine();
                Console.CursorVisible = false;

                Console.WriteLine("Please enter html file path:");

                Console.CursorVisible = true;
                htmlFileName = Console.ReadLine();
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
                    PageTitle = Path.GetFileName(xlsxFileName)
                };

                //Convert the Xlsx file
                using (FileStream outputStream = new FileStream(htmlFileName, FileMode.Create))
                {
                    XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, outputStream, config, progressCallback);
                }

                //Open the Html file
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(htmlFileName) { UseShellExecute = true, CreateNoWindow = true });
                }
                catch
                {
                    Console.WriteLine("\nPress Enter key to exit.");
                    Console.ReadLine();
                }
                finally
                {
                    Environment.Exit(1);
                }
            }
            catch (Exception ex)
            {
                //Output the error
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("\nPress Enter key to exit.");
                Console.ReadLine();

                Environment.Exit(0);
            }
        }

        private static void ConverterProgressCallback(object sender, XlsxToHtmlConverter.ConverterProgressCallbackEventArgs e)
        {
            //Output the progress
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.WriteLine(string.Format("{0:##0.00}% (Sheet {1} of {2} | Row {3} of {4})", e.ProgressPercent, e.CurrentSheet, e.TotalSheets, e.CurrentRow, e.TotalRows) + new string(' ', 5) + new string('█', (int)(e.ProgressPercent / 2)) + new string('░', (int)((100 - e.ProgressPercent) / 2)));
        }
    }
}
