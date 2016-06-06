//-----------------------------------------------------------------------------------------------------------
// <copyright file="ProgressReport.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfConversionAndTimeStampTool
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