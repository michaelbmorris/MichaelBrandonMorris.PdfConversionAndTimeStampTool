//-----------------------------------------------------------------------------------------------------------
// <copyright file="Script.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfConversionAndTimeStampTool
{
    using static Properties.Resources;

    internal enum ScriptEvent
    {
        WillPrint,
        WillSave,
        DidPrint,
        DidSave
    }

    internal class Script
    {
        internal static readonly Script TimeStampOnPrintDefaultDayScript =
            new Script(TimeStampOnPrintDefaultDay, ScriptEvent.WillPrint);

        internal static readonly Script TimeStampOnPrintDefaultMonthScript =
            new Script(TimeStampOnPrintDefaultMonth, ScriptEvent.WillPrint);

        internal Script(string scriptText, ScriptEvent scriptEvent)
        {
            ScriptText = scriptText;
            ScriptEvent = scriptEvent;
        }

        internal string ScriptText { get; set; }

        internal ScriptEvent ScriptEvent { get; set; }
    }
}