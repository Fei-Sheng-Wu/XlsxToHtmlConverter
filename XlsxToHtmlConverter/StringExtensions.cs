using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace XlsxToHtmlConverter
{
    public static class StringExtensions
    {
        public static string ToInvariant(this double value)
        {
            // Convert the double to a string using InvariantCulture
            return value.ToString(CultureInfo.InvariantCulture);
        }
    }
}
