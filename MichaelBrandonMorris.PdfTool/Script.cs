using System.Collections.Generic;
using static MichaelBrandonMorris.PdfTool.Properties.Resources;

namespace MichaelBrandonMorris.PdfTool
{
    /// <summary>
    ///     Enum ScriptTiming
    /// </summary>
    /// TODO Edit XML Comment Template for ScriptTiming
    internal enum ScriptTiming
    {
        /// <summary>
        ///     The did print
        /// </summary>
        /// TODO Edit XML Comment Template for DidPrint
        DidPrint,

        /// <summary>
        ///     The did save
        /// </summary>
        /// TODO Edit XML Comment Template for DidSave
        DidSave,

        /// <summary>
        ///     The will print
        /// </summary>
        /// TODO Edit XML Comment Template for WillPrint
        WillPrint,

        /// <summary>
        ///     The will save
        /// </summary>
        /// TODO Edit XML Comment Template for WillSave
        WillSave
    }

    /// <summary>
    ///     Class Script.
    /// </summary>
    /// TODO Edit XML Comment Template for Script
    internal class Script
    {
        /// <summary>
        ///     The time stamp on print day
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampOnPrintDay
        internal static readonly Script TimeStampOnPrintDay =
            new Script(TimeStampDay, ScriptTiming.WillPrint);

        /// <summary>
        ///     The time stamp on print month
        /// </summary>
        /// TODO Edit XML Comment Template for TimeStampOnPrintMonth
        internal static readonly Script TimeStampOnPrintMonth =
            new Script(TimeStampMonth, ScriptTiming.WillPrint);

        /// <summary>
        ///     Initializes a new instance of the <see cref="Script" />
        ///     class.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="timing">The timing.</param>
        /// TODO Edit XML Comment Template for #ctor
        internal Script(string text, ScriptTiming timing)
        {
            Text = text;
            Timing = timing;
        }

        /// <summary>
        ///     Gets the text.
        /// </summary>
        /// <value>The text.</value>
        /// TODO Edit XML Comment Template for Text
        internal string Text
        {
            get;
        }

        /// <summary>
        ///     Gets the timing.
        /// </summary>
        /// <value>The timing.</value>
        /// TODO Edit XML Comment Template for Timing
        internal ScriptTiming Timing
        {
            get;
        }

        /// <summary>
        ///     Gets the script timings.
        /// </summary>
        /// <returns>IList&lt;ScriptTiming&gt;.</returns>
        /// TODO Edit XML Comment Template for GetScriptTimings
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