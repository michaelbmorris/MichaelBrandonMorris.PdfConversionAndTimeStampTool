namespace PdfScriptTool
{
    internal class Field
    {
        private const int XMax = 612;
        private const int TYMax = 792;
        private const int XMin = 0;
        private const int YMin = 0;
        private const int DefaultTopLeftX = 36;
        private const int DefaultTopLeftY = 792;
        private const int DefaultBottomRightX = 576;
        private const int DefaultBottomRightY = 756;

        public int TopLeftX
        {
            get
            {
                return coordinates[0];
            }
        }

        public int TopLeftY
        {
            get
            {
                return coordinates[1];
            }
        }

        public int BottomRightX
        {
            get
            {
                return coordinates[2];
            }
        }

        public int BottomRightY
        {
            get
            {
                return coordinates[3];
            }
        }

        public string Title { get; set; }

        private int[] coordinates;

        public Field(string title, int topLeftX, int topLeftY, int bottomRightX, int bottomRightY)
        {
            Title = title;
            coordinates = new int[4];
            coordinates[0] = topLeftX;
            coordinates[1] = topLeftY;
            coordinates[2] = bottomRightX;
            coordinates[3] = bottomRightY;
        }

        public static readonly Field DefaultField
            = new Field(
                "Timestamp",
                DefaultTopLeftX,
                DefaultTopLeftY,
                DefaultBottomRightX,
                DefaultBottomRightY);
    }
}