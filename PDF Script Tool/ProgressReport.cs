//-----------------------------------------------------------------------------------------------------------
// <copyright file="Script.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    internal class ProgressReport
    {
        private const int PercentMultiplier = 100;
        internal int CurrentCount { get; set; }

        internal int Percent
        {
            get
            {
                return PercentMultiplier * CurrentCount / Total;
            }
        }

        internal int Total { get; set; }
    }
}