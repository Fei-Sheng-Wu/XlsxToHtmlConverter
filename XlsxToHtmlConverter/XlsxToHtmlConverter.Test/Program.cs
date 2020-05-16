using System;
using System.Linq;

namespace XlsxToHtmlConverter.Test
{
    class Program
    {
        static void Main()
        {
            try
            {
                Console.WriteLine("Please enter xlsx file path:");
                string xlsxFileName = Console.ReadLine();
                Console.WriteLine("Please enter html file path:");
                string htmlFileName = Console.ReadLine();

                Console.WriteLine("\nConverting...");

                DateTime startTime = DateTime.UtcNow;

                string htmlString = XlsxToHtmlConverter.Converter.ConvertXlsx(xlsxFileName);

                System.IO.File.WriteAllText(htmlFileName, htmlString);

                DateTime endTime = DateTime.UtcNow;

                Console.WriteLine("\nConvert end. Total seconds: " + (endTime - startTime).TotalSeconds + ".");
                Console.WriteLine("\nPress Enter key to continue.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError: " + ex.Message);
                Console.WriteLine("\nPress Enter key to continue.");
                Console.ReadLine();
            }
        }
    }
}
