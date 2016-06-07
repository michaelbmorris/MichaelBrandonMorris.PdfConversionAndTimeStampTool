//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfConversionAndTimeStampTool.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfConversionAndTimeStampTool
{
    using System.IO;
    using Application = System.Windows.Forms.Application;

    internal static class PdfConversionAndTimeStampTool
    {
        [System.STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var view = new PdfConversionAndTimeStampToolView();
            var presenter = new PdfConversionAndTimeStampToolPresenter(view);
            Directory.CreateDirectory(FileProcessor.OutputPath);
            Directory.CreateDirectory(FileProcessor.ProcessingPath);
            FileProcessor.ClearProcessing();
            Application.Run(view);
        }
    }
}