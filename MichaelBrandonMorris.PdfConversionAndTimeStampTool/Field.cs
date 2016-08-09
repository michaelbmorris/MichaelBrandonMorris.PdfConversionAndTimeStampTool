using System.Collections.Generic;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.FieldPages;

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
            TimeStampFieldName,
            TimeStampFieldLeftX,
            TimeStampFieldTopY,
            TimeStampFieldRightX,
            TimeStampFieldBottomY,
            All);

        private const int TimeStampFieldBottomY = 756;
        private const int TimeStampFieldLeftX = 36;
        private const string TimeStampFieldName = "Timestamp";
        private const int TimeStampFieldRightX = 576;
        private const int TimeStampFieldTopY = 792;

        internal Field(
            string name,
            int leftX,
            int topY,
            int rightX,
            int bottomY,
            FieldPages pages,
            IEnumerable<int> customPageNumbers = null)
        {
            Name = name;
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

        internal IEnumerable<int> CustomPageNumbers
        {
            get;
        }

        internal int LeftX
        {
            get;
        }

        internal string Name
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

        internal int TopY
        {
            get;
        }

        internal static IList<FieldPages> GetFieldPages()
        {
            return new List<FieldPages>
            {
                All,
                Custom,
                Even,
                First,
                Last,
                Odd
            };
        }
    }
}