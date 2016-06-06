//-----------------------------------------------------------------------------------------------------------
// <copyright file="Field.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfConversionAndTimeStampTool
{
    using static Properties.Resources;

    internal enum Pages
    {
        All,
        Odd,
        Even,
        First,
        Last
    }

    internal class Field
    {
        internal static readonly Field DefaultTimeStampField
            = new Field(
                DefaultTimestampFieldTitle,
                DefaultTopLeftX,
                DefaultTopLeftY,
                DefaultBottomRightX,
                DefaultBottomRightY,
                Pages.All);

        private const int DefaultBottomRightX = 576;

        private const int DefaultBottomRightY = 756;

        private const int DefaultTopLeftX = 36;

        private const int DefaultTopLeftY = 792;

        private const int TYMax = 792;

        private const int XMax = 612;

        private const int XMin = 0;

        private const int YMin = 0;

        private int[] coordinates;

        internal Field(
            string title,
            int topLeftX,
            int topLeftY,
            int bottomRightX,
            int bottomRightY,
            Pages pages)
        {
            Title = title;
            coordinates = new int[4];
            coordinates[0] = topLeftX;
            coordinates[1] = topLeftY;
            coordinates[2] = bottomRightX;
            coordinates[3] = bottomRightY;
            Pages = pages;
        }

        internal int BottomRightX
        {
            get
            {
                return coordinates[2];
            }
        }

        internal int BottomRightY
        {
            get
            {
                return coordinates[3];
            }
        }

        internal Pages Pages { get; set; }

        internal string Title { get; set; }

        internal int TopLeftX
        {
            get
            {
                return coordinates[0];
            }
        }

        internal int TopLeftY
        {
            get
            {
                return coordinates[1];
            }
        }
    }
}