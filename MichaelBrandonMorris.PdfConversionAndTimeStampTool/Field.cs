using System.Collections.Generic;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    internal enum FieldPages
    {
        All,
        Custom,
        Even,
        First,
        Last,
        Odd
    }

    internal class Field
    {
        internal static readonly Field TimeStampField = new Field(
            "Timestamp", 36, 792, 576, 756, FieldPages.All);

        internal Field(
            string title,
            int leftX,
            int topY,
            int rightX,
            int bottomY,
            FieldPages pages,
            IEnumerable<int> customPageNumbers = null)
        {
            Title = title;
            LeftX = leftX;
            TopY = topY;
            RightX = rightX;
            BottomY = bottomY;
            Pages = pages;
            CustomPageNumbers = customPageNumbers;
        }

        internal int BottomY
        {
            get;
        }

        internal int LeftX
        {
            get;
        }

        internal FieldPages Pages
        {
            get;
        }

        internal int RightX
        {
            get;
        }

        internal string Title
        {
            get;
        }

        internal int TopY
        {
            get;
        }

        private IEnumerable<int> CustomPageNumbers
        {
            get;
        }

        internal static IList<FieldPages> GetFieldPages()
        {
            return new List<FieldPages>
            {
                FieldPages.All,
                FieldPages.Custom,
                FieldPages.Even,
                FieldPages.First,
                FieldPages.Last,
                FieldPages.Odd
            };
        }
    }
}