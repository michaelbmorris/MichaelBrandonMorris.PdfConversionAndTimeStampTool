//-----------------------------------------------------------------------------------------------------------
// <copyright file="Script.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    using static Properties.Resources;

    /// <summary>
    /// The event that triggers the execution of the script.
    /// </summary>
    internal enum ScriptEvent
    {
        /// <summary>
        /// The document is preparing to print.
        /// </summary>
        WillPrint,

        /// <summary>
        /// The document is preparing to save.
        /// </summary>
        WillSave,

        /// <summary>
        /// The document was printed.
        /// </summary>
        DidPrint,

        /// <summary>
        /// The document was saved.
        /// </summary>
        DidSave
    }

    /// <summary>
    /// Scripts define JavaScript text that can be added to a PDF file.
    /// </summary>
    internal class Script
    {
        /// <summary>
        /// The default "time stamp on print" script, valid for a day.
        /// </summary>
        internal static readonly Script TimeStampOnPrintDefaultDayScript =
            new Script(TimeStampOnPrintDefaultDay, ScriptEvent.WillPrint);

        /// <summary>
        /// The default "time stamp on print" script, valid for a month.
        /// </summary>
        internal static readonly Script TimeStampOnPrintDefaultMonthScript =
            new Script(TimeStampOnPrintDefaultMonth, ScriptEvent.WillPrint);

        /// <summary>
        /// Initializes a new instance of the <see cref="Script"/> class.
        /// </summary>
        /// <param name="scriptText">The JavaScript text of the script.</param>
        /// <param name="scriptEvent">The event that triggers the execution of
        /// the script.</param>
        internal Script(string scriptText, ScriptEvent scriptEvent)
        {
            this.ScriptText = scriptText;
            this.ScriptEvent = scriptEvent;
        }

        /// <summary>
        /// Gets or sets the JavaScript text of the script.
        /// </summary>
        internal string ScriptText { get; set; }

        /// <summary>
        /// Gets or sets the event that triggers the execution of the script.
        /// </summary>
        internal ScriptEvent ScriptEvent { get; set; }
    }
}