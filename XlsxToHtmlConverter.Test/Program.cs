using System;

namespace XlsxToHtmlConverter.Test
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.CursorVisible = false;

            string xlsxFileName;
            string htmlFileName;

            //Get input and output file paths
            if (args != null && args.Length == 2)
            {
                xlsxFileName = args[0];
                htmlFileName = args[1];
            }
            else if (args != null && args.Length == 1)
            {
                xlsxFileName = args[0];
                htmlFileName = System.IO.Path.ChangeExtension(xlsxFileName, "html");
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
                //Create progress callback event
                EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> converterProgressCallbackEvent = null;
                converterProgressCallbackEvent += ConverterProgressCallback;

                //Convert document
                string htmlString = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, converterProgressCallbackEvent);
                System.IO.File.WriteAllText(htmlFileName, htmlString);

                //Open html file
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
                    //Exit
                    Environment.Exit(1);
                }
            }
            catch (Exception ex)
            {
                //Output error
                Console.SetCursorPosition(0, Console.CursorTop - 1);
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("\nPress Enter key to exit.");
                Console.ReadLine();

                //Exit
                Environment.Exit(0);
            }
        }

        private static void ConverterProgressCallback(object sender, ConverterProgressCallbackEventArgs e)
        {
            //Output the progress
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.WriteLine(String.Format("{0:##0.00}% ({1} of {2})", e.ProgressPercent, e.CurrentSheet, e.TotalSheets).PadRight(String.Format("100.00% ({0} of {0})", e.TotalSheets).Length + 1) + new string('█', (int)(e.ProgressPercent / 2)) + new string('░', (int)((100 - e.ProgressPercent) / 2)));
        }
    }
}
