using System;
using System.Collections.Generic;
using System.IO;

namespace PdfConversionAndTimeStampTool
{
    internal static class StringExtensions
    {
        internal static bool FileNameEquals(
            this string arg1, string arg2)
        {
            return string.Equals(
                Path.GetFileNameWithoutExtension(arg1),
                Path.GetFileNameWithoutExtension(arg2),
                StringComparison.InvariantCultureIgnoreCase);
        }

        internal static bool FileNameIsContainedIn(
            this string arg1, IEnumerable<string> arg2)
        {
            foreach (string s in arg2)
            {
                if (arg1.FileNameEquals(s))
                {
                    return true;
                }
            }
            return false;
        }
    }
}