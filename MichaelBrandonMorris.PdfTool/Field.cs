using System.Collections.Generic;

namespace MichaelBrandonMorris.PdfTool
{
    /// <summary>
    ///     Enum FieldPages
    /// </summary>
    /// TODO Edit XML Comment Template for FieldPages
    internal enum FieldPages
    {
        /// <summary>
        ///     All
        /// </summary>
        /// TODO Edit XML Comment Template for All
        All,

        /// <summary>
        ///     The custom
        /// </summary>
        /// TODO Edit XML Comment Template for Custom
        Custom,

        /// <summary>
        ///     The even
        /// </summary>
        /// TODO Edit XML Comment Template for Even
        Even,

        /// <summary>
        ///     The first
        /// </summary>
        /// TODO Edit XML Comment Template for First
        First,

        /// <summary>
        ///     The last
        /// </summary>
        /// TODO Edit XML Comment Template for Last
        Last,

        /// <summary>
        ///     The odd
        /// </summary>
        /// TODO Edit XML Comment Template for Odd
        Odd
    }

    /// <summary>
    ///     Class Field.
    /// </summary>
    /// TODO Edit XML Comment Template for Field
    internal class Field
    {
        /// <summary>
        ///     The time stamp field
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampField
        internal static readonly Field TimeStampField = new Field(
            TimeStampFieldName,
            TimeStampFieldLeftX,
            TimeStampFieldTopY,
            TimeStampFieldRightX,
            TimeStampFieldBottomY,
            FieldPages.All);

        /// <summary>
        ///     The time stamp field bottom y
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampFieldBottomY
        private const int TimeStampFieldBottomY = 756;

        /// <summary>
        ///     The time stamp field left x
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampFieldLeftX
        private const int TimeStampFieldLeftX = 36;

        /// <summary>
        ///     The time stamp field name
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampFieldName
        private const string TimeStampFieldName = "Timestamp";

        /// <summary>
        ///     The time stamp field right x
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampFieldRightX
        private const int TimeStampFieldRightX = 576;

        /// <summary>
        ///     The time stamp field top y
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampFieldTopY
        private const int TimeStampFieldTopY = 792;

        /// <summary>
        ///     Initializes a new instance of the <see cref="Field" />
        ///     class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="leftX">The left x.</param>
        /// <param name="topY">The top y.</param>
        /// <param name="rightX">The right x.</param>
        /// <param name="bottomY">The bottom y.</param>
        /// <param name="pages">The pages.</param>
        /// <param name="customPageNumbers">The custom page numbers.</param>
        /// TODO Edit XML Comment Template for #ctor
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

        /// <summary>
        ///     Gets the bottom y.
        /// </summary>
        /// <value>The bottom y.</value>
        /// TODO Edit XML Comment Template for BottomY
        internal int BottomY
        {
            get;
        }

        /// <summary>
        ///     Gets the custom page numbers.
        /// </summary>
        /// <value>The custom page numbers.</value>
        /// TODO Edit XML Comment Template for CustomPageNumbers
        internal IEnumerable<int> CustomPageNumbers
        {
            get;
        }

        /// <summary>
        ///     Gets the left x.
        /// </summary>
        /// <value>The left x.</value>
        /// TODO Edit XML Comment Template for LeftX
        internal int LeftX
        {
            get;
        }

        /// <summary>
        ///     Gets the name.
        /// </summary>
        /// <value>The name.</value>
        /// TODO Edit XML Comment Template for Name
        internal string Name
        {
            get;
        }

        /// <summary>
        ///     Gets the pages.
        /// </summary>
        /// <value>The pages.</value>
        /// TODO Edit XML Comment Template for Pages
        internal FieldPages Pages
        {
            get;
        }

        /// <summary>
        ///     Gets the right x.
        /// </summary>
        /// <value>The right x.</value>
        /// TODO Edit XML Comment Template for RightX
        internal int RightX
        {
            get;
        }

        /// <summary>
        ///     Gets the top y.
        /// </summary>
        /// <value>The top y.</value>
        /// TODO Edit XML Comment Template for TopY
        internal int TopY
        {
            get;
        }

        /// <summary>
        ///     Gets the field pages.
        /// </summary>
        /// <returns>IList&lt;FieldPages&gt;.</returns>
        /// TODO Edit XML Comment Template for GetFieldPages
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