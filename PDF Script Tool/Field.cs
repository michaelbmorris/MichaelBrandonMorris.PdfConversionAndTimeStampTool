//-----------------------------------------------------------------------------------------------------------
// <copyright file="Field.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    using static Properties.Resources;

    /// <summary>
    /// The allowed specifications for which pages a field can be placed on.
    /// </summary>
    internal enum Pages
    {
        /// <summary>
        /// All pages in the document.
        /// </summary>
        All,

        /// <summary>
        /// Odd pages in the document.
        /// </summary>
        Odd,

        /// <summary>
        /// Even pages in the document.
        /// </summary>
        Even,

        /// <summary>
        /// The first page of the document.
        /// </summary>
        First,

        /// <summary>
        /// The last page of the document.
        /// </summary>
        Last
    }

    /// <summary>
    /// Fields are used to design text fields for PDF files.
    /// </summary>
    internal class Field
    {
        /// <summary>
        /// The default field used for "time stamp on print" methods.
        /// </summary>
        internal static readonly Field DefaultTimeStampField
            = new Field(
                DefaultTimestampFieldTitle,
                DefaultTopLeftX,
                DefaultTopLeftY,
                DefaultBottomRightX,
                DefaultBottomRightY,
                Pages.All);

        /// <summary>
        /// The x coordinate of the default bottom right field corner.
        /// </summary>
        private const int DefaultBottomRightX = 576;

        /// <summary>
        /// The y coordinate of the default bottom right field corner.
        /// </summary>
        private const int DefaultBottomRightY = 756;

        /// <summary>
        /// The x coordinate of the default top left field corner.
        /// </summary>
        private const int DefaultTopLeftX = 36;

        /// <summary>
        /// The y coordinate of the default top left field corner.
        /// </summary>
        private const int DefaultTopLeftY = 792;

        /// <summary>
        /// The maximum y coordinate in points on a portrait-oriented PDF.
        /// Top edge of document.
        /// </summary>
        private const int TYMax = 792;

        /// <summary>
        /// The maximum x coordinate in points on a portrait-oriented PDF.
        /// Right edge of document.
        /// </summary>
        private const int XMax = 612;

        /// <summary>
        /// The minimum x coordinate in points on a PDF.
        /// Left edge of document.
        /// </summary>
        private const int XMin = 0;

        /// <summary>
        /// The minimum y coordinate in points on a PDF.
        /// Bottom edge of document.
        /// </summary>
        private const int YMin = 0;

        /// <summary>
        /// The four coordinates that comprise the top left and bottom right
        /// corners of the field.
        /// </summary>
        private int[] coordinates;

        /// <summary>
        /// Initializes a new instance of the <see cref="Field" /> class.
        /// </summary>
        /// <param name="title">
        /// The title of the field.
        /// </param>
        /// <param name="topLeftX">
        /// The top left corner x coordinate of the field.
        /// </param>
        /// <param name="topLeftY">
        /// The top left corner y coordinate of the field.
        /// </param>
        /// <param name="bottomRightX">
        /// The bottom right x coordinate of the field.
        /// </param>
        /// <param name="bottomRightY">
        /// The bottom right y coordinate of the field.
        /// </param>
        /// <param name="pages">The pages the field should be placed on.
        /// </param>
        internal Field(
            string title,
            int topLeftX,
            int topLeftY,
            int bottomRightX,
            int bottomRightY,
            Pages pages)
        {
            this.Title = title;
            this.coordinates = new int[4];
            this.coordinates[0] = topLeftX;
            this.coordinates[1] = topLeftY;
            this.coordinates[2] = bottomRightX;
            this.coordinates[3] = bottomRightY;
            this.Pages = pages;
        }

        /// <summary>
        /// Gets the bottom right corner x coordinate of the field.
        /// </summary>
        internal int BottomRightX
        {
            get
            {
                return this.coordinates[2];
            }
        }

        /// <summary>
        /// Gets the bottom right corner y coordinate of the field.
        /// </summary>
        internal int BottomRightY
        {
            get
            {
                return this.coordinates[3];
            }
        }

        /// <summary>
        /// Gets or sets the pages the field should be placed on.
        /// </summary>
        internal Pages Pages { get; set; }

        /// <summary>
        /// Gets or sets the title of the field.
        /// </summary>
        internal string Title { get; set; }

        /// <summary>
        /// Gets the top left corner x coordinate of the field.
        /// </summary>
        internal int TopLeftX
        {
            get
            {
                return this.coordinates[0];
            }
        }

        /// <summary>
        /// Gets the top left corner y coordinate of the field.
        /// </summary>
        internal int TopLeftY
        {
            get
            {
                return this.coordinates[1];
            }
        }
    }
}