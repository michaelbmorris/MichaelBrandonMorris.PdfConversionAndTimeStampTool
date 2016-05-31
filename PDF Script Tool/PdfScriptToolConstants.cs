namespace PdfScriptTool
{
    internal static class PdfScriptToolConstants
    {
        internal const string DefaultTimeStampScript
            = "﻿var f = this.getField('Timestamp');"
            + "f.alignment = 'left';"
            + "f.multiline = false;"
            + "f.textSize = 11;"
            + "f.richText = true;"
            + "var style = f.defaultStyle;"
            + "style.fontFamily = ['Calibri', 'sans-serif'];"
            + "f.defaultStyle = style;"
            + "var t = new Array();"
            + "t[0] = new Object();"
            + "t[0].text = 'Uncontrolled 24 hours after ';"
            + "t[1] = new Object();"
            + "t[1].text = util.printd('mm/dd/yy h:MM tt', new Date());"
            + "f.richValue = t;";

        internal const string RootFolderName = "PDFScriptTool";
        internal const string OutputFolderName = "Output";
        internal const string ConfigurationFolderName = "Configuration";
        internal const string ProcessingFolderName = "Processing";

        internal const string OpenFileDialogFilter
            = "Documents (*.doc;*.docx;*.pdf)|*.doc;*.docx;*.pdf";

        internal const string OpenFileDialogTitle = "Select documents...";
        internal const bool OpenFileDialogAllowMultiple = true;
        internal const bool DocumentsViewFileIsChecked = true;
        internal const string TimeStampFieldName = "Timestamp";
        internal const int PdfFirstPageNumber = 1;

        internal const int TimeStampFieldUnderlineLeftXCoordinate = 36;
        internal const int TimeStampFieldUnderlineRightXCoordinate = 576;
        internal const int TimeStampFieldUnderlineYCoordinate = 768;
        internal const int PercentMultiplier = 100;
        internal const string ProgressLabelDivider = " of ";
    }
}