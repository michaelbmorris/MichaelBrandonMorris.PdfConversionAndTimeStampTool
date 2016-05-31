using PdfScriptTool.Properties;

namespace PdfScriptTool
{
    internal enum ScriptTiming
    {
        WillPrint,
        WillSave,
        DidPrint,
        DidSave
    }

    internal class Script
    {
        public string Text { get; set; }
        public ScriptTiming Timing { get; set; }
        public Field Field { get; set; }

        internal Script(string scriptText, ScriptTiming timing, Field field)
        {
            Text = scriptText;
            Timing = timing;
            Field = field;
        }

        public static readonly Script TimeStampOnPrintDefaultDay = new Script(
            Resources.TimeStampOnPrintDefaultDay,
            ScriptTiming.WillPrint,
            Field.DefaultField);

        public static readonly Script TimeStampOnPrintDefaultMonth = new Script(
            Resources.TimeStampOnPrintDefaultMonth,
            ScriptTiming.WillPrint,
            Field.DefaultField);
    }
}