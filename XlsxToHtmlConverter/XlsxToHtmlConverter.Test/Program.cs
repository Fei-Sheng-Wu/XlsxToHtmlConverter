using System;

namespace XlsxToHtmlConverter.Test
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                //Get input and output file paths
                Console.WriteLine("Please enter xlsx file path:");
                string xlsxFileName = Console.ReadLine();
                Console.WriteLine("Please enter html file path:");
                string htmlFileName = Console.ReadLine();

                //Start convert
                Console.WriteLine("\nConverting...");
                Console.WriteLine();
                DateTime startTime = DateTime.UtcNow;

                //Create progress callback event
                EventHandler<XlsxToHtmlConverter.ConverterProgressCallbackEventArgs> converterProgressCallbackEvent = null;
                converterProgressCallbackEvent += ConverterProgressCallback;

                //Convert document
                string htmlString = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName, converterProgressCallbackEvent);
                System.IO.File.WriteAllText(htmlFileName, htmlString);

                //End convert
                DateTime endTime = DateTime.UtcNow;
                Console.WriteLine("\nConvert end. Total seconds: " + (endTime - startTime).TotalSeconds + ".");
                Console.WriteLine("\nPress Enter key to continue.");
                Console.ReadLine();

                //Exit
                Environment.Exit(1);
            }
            catch (Exception ex)
            {
                //Output error
                Console.WriteLine("\nError: " + ex.Message);
                Console.WriteLine("\nPress Enter key to continue.");
                Console.ReadLine();

                //Exit
                Environment.Exit(0);
            }
        }

        private static void ConverterProgressCallback(object sender, ConverterProgressCallbackEventArgs e)
        {
            //Output the progress
            Console.WriteLine(String.Format("{0:##0.00}% ({1} of {2})", e.ProgressPercent, e.CurrentSheet, e.TotalSheets).PadRight(String.Format("100.00% ({0} of {0})", e.TotalSheets).Length + 1) + new string('*', (int)(e.ProgressPercent / 2)));
        }
    }
}
