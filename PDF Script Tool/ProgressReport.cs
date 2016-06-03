//-----------------------------------------------------------------------------------------------------------
// <copyright file="ProgressReport.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfTool
{
    /// <summary>
    /// Allows for reporting progress as both a percent and a count.
    /// </summary>
    internal class ProgressReport
    {
        /// <summary>
        /// Converts a decimal number to a percent value.
        /// </summary>
        private const int PercentMultiplier = 100;

        /// <summary>
        /// Gets or sets the number of tasks completed.
        /// </summary>
        internal int CurrentCount { get; set; }

        /// <summary>
        /// Gets the percent of tasks completed.
        /// </summary>
        internal int Percent
        {
            get
            {
                return PercentMultiplier * CurrentCount / Total;
            }
        }

        /// <summary>
        /// Gets or sets the total number of tasks.
        /// </summary>
        internal int Total { get; set; }
    }
}