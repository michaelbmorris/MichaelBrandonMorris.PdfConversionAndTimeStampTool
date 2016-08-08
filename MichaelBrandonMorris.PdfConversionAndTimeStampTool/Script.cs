using System.Collections.Generic;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.Properties.Resources;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    internal enum ScriptTiming
    {
        DidPrint,
        DidSave,
        WillPrint,
        WillSave
    }

    internal class Script
    {
        internal static readonly Script TimeStampOnPrintDay = new Script(
            TimeStampDay, ScriptTiming.WillPrint);

        internal static readonly Script TimeStampOnPrintMonth = new Script(
            TimeStampMonth, ScriptTiming.WillPrint);

        internal Script(string text, ScriptTiming timing)
        {
            Text = text;
            Timing = timing;
        }

        internal string Text
        {
            get;
        }

        internal ScriptTiming Timing
        {
            get;
        }

        internal static IList<ScriptTiming> GetScriptTimings()
        {
            return new List<ScriptTiming>
            {
                ScriptTiming.DidPrint,
                ScriptTiming.DidSave,
                ScriptTiming.WillPrint,
                ScriptTiming.WillSave
            };
        }
    }
}