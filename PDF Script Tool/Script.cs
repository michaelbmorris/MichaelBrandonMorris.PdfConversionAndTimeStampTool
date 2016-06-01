//-----------------------------------------------------------------------------------------------------------
// <copyright file="Script.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
    using static Properties.Resources;

    internal enum ScriptTiming
    {
        WillPrint,
        WillSave,
        DidPrint,
        DidSave
    }

    internal class Script
    {
        public string ScriptText { get; set; }
        public ScriptTiming ScriptTiming { get; set; }

        internal Script(string scriptText, ScriptTiming scriptTiming)
        {
            ScriptText = scriptText;
            ScriptTiming = scriptTiming;
        }

        public static readonly Script TimeStampOnPrintDefaultDayScript = new Script(
            TimeStampOnPrintDefaultDay,
            ScriptTiming.WillPrint);

        public static readonly Script TimeStampOnPrintDefaultMonthScript = new Script(
            TimeStampOnPrintDefaultMonth,
            ScriptTiming.WillPrint);
    }
}