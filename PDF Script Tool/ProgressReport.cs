using static PdfScriptTool.PdfScriptToolConstants;

namespace PdfScriptTool
{
    internal class ProgressReport
    {
        public int CurrentCount { get; set; }

        public int Percent
        {
            get
            {
                return PercentMultiplier * CurrentCount / Total;
            }
        }

        public int Total { get; set; }
    }
}